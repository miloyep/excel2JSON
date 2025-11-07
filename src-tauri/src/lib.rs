use calamine::{open_workbook, DataType, Reader, Xlsx};
use chrono::{Duration as ChronoDuration, Local, NaiveDate};
use indexmap::IndexMap;
use json::JsonValue;
use serde::Serialize;
use serde_json::json;
use std::fs::{self, File};
use std::io::BufReader;
use std::path::{Path, PathBuf};
use tauri::{AppHandle, Emitter};
use walkdir::WalkDir;
use zip::write::FileOptions;
use zip::CompressionMethod;

/// 日志类型
#[derive(Serialize)]
#[serde(rename_all = "lowercase")]
enum LogType {
    Info,
    Success,
    Warning,
    Error,
}

/// 封装事件发送方法，同时包含消息类型
fn send_progress(app: &AppHandle, msg: &str, log_type: LogType) -> Result<(), String> {
    let payload = json!({
        "message": msg,
        "type": log_type
    });
    app.emit("excel-export-progress", payload)
        .map_err(|e| format!("发送进度事件失败: {}", e))
}

/// 检查字符串里的 {{}} 是否完整
fn check_placeholders(value: &str) -> Result<(), String> {
    let mut stack = 0;
    for ch in value.chars() {
        match ch {
            '{' => stack += 1,
            '}' => {
                if stack == 0 {
                    return Err(format!("占位符不完整: {}", value));
                }
                stack -= 1;
            }
            _ => {}
        }
    }
    if stack != 0 {
        return Err(format!("占位符不完整: {}", value));
    }
    Ok(())
}

// 配置结构体
#[derive(Debug, Clone)]
struct SheetConfig {
    name: String,
    sheet_type: Option<String>,
}

#[derive(Debug, Clone)]
struct LanguageConfig {
    code: String,
}

// 从 Excel 读取语言配置
fn read_language_configs_from_excel(
    workbook: &mut Xlsx<BufReader<fs::File>>,
) -> Result<Vec<LanguageConfig>, Box<dyn std::error::Error>> {
    let sheet_name = "导出语言管理";
    let range = workbook
        .worksheet_range(sheet_name)
        .ok_or_else(|| format!("未找到工作表: {}", sheet_name))??;

    let mut configs = Vec::new();

    for row in range.rows() {
        if !row.is_empty() {
            let lang_code = get_cell_string(&row[0]);
            if !lang_code.is_empty() {
                configs.push(LanguageConfig { code: lang_code });
            }
        }
    }

    Ok(configs)
}

fn read_sheet_configs_from_excel(
    app: &AppHandle,
    workbook: &mut Xlsx<BufReader<fs::File>>,
) -> Result<Vec<SheetConfig>, Box<dyn std::error::Error>> {
    let sheet_name = "导出sheet管理";
    let range = workbook
        .worksheet_range(sheet_name)
        .ok_or_else(|| format!("未找到工作表: {}", sheet_name))??;

    let mut configs = Vec::new();

    for row in range.rows() {
        if row.len() < 1 {
            continue;
        }

        let name = get_cell_string(&row[0]);
        let sheet_type = if row.len() > 1 {
            let s = get_cell_string(&row[1]);
            if s.trim().is_empty() {
                None
            } else {
                Some(s)
            }
        } else {
            None
        };

        configs.push(SheetConfig { name, sheet_type });
    }

    Ok(configs)
}

fn get_cell_string(cell: &DataType) -> String {
    match cell {
        DataType::String(s) => s.to_string(),
        DataType::Float(f) => {
            if *f > 30000.0 && *f < 70000.0 {
                let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                let date = base_date + ChronoDuration::days(*f as i64);
                date.format("%Y-%m-%d").to_string()
            } else if f.fract() == 0.0 {
                format!("{:.0}", f)
            } else {
                f.to_string()
            }
        }
        DataType::Int(i) => i.to_string(),
        DataType::Bool(b) => b.to_string(),
        DataType::DateTime(dt) => {
            let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
            let days = dt.trunc() as i64;
            let seconds = ((*dt - dt.trunc()) * 86400.0) as i64;
            let datetime = base_date.and_hms_opt(0, 0, 0).unwrap()
                + ChronoDuration::days(days)
                + ChronoDuration::seconds(seconds);
            datetime.format("%Y-%m-%d %H:%M:%S").to_string()
        }
        DataType::Duration(d) => {
            let total_seconds = (*d * 86400.0) as i64;
            let hours = total_seconds / 3600;
            let minutes = (total_seconds % 3600) / 60;
            let seconds = total_seconds % 60;
            format!("{:02}:{:02}:{:02}", hours, minutes, seconds)
        }
        DataType::DateTimeIso(s) => s.clone(),
        DataType::DurationIso(s) => s.clone(),
        DataType::Error(e) => format!("Error: {:?}", e),
        DataType::Empty => String::new(),
    }
}

/// 压缩整个文件夹为 zip 文件
fn zip_directory(src_dir: &Path, dst_file: &Path) -> Result<(), String> {
    let file = File::create(dst_file).map_err(|e| format!("创建 zip 文件失败: {}", e))?;
    let mut zip = zip::ZipWriter::new(file);
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

    let base_path = src_dir.parent().unwrap_or_else(|| Path::new(""));

    for entry in WalkDir::new(src_dir) {
        let entry = entry.map_err(|e| format!("读取目录失败: {}", e))?;
        let path = entry.path();

        if path.is_file() {
            // 计算相对路径（去掉上级目录）
            let name = path.strip_prefix(base_path).unwrap();
            let name_str = name.to_string_lossy();

            let mut f = File::open(path).map_err(|e| format!("打开文件失败: {}", e))?;
            zip.start_file(name_str, options)
                .map_err(|e| format!("写入 zip 条目失败: {}", e))?;

            std::io::copy(&mut f, &mut zip).map_err(|e| format!("写入 zip 内容失败: {}", e))?;
        }
    }

    zip.finish().map_err(|e| format!("关闭 zip 失败: {}", e))?;
    Ok(())
}

#[tauri::command]
async fn convert_excel_to_json(app: AppHandle, path: String) -> Result<String, String> {
    send_progress(&app, &format!("开始处理文件: {}", path), LogType::Info)?;

    let file_path = PathBuf::from(&path);
    if !file_path.exists() {
        let msg = format!("文件不存在: {}", file_path.display());
        send_progress(&app, &msg, LogType::Error)?;
        return Err(msg);
    }

    send_progress(&app, "正在打开 Excel 文件...", LogType::Info)?;
    let mut workbook: Xlsx<_> =
        open_workbook(&file_path).map_err(|e| format!("打开文件失败: {}", e))?;
    send_progress(&app, "Excel 文件已成功打开", LogType::Success)?;

    let lang_configs = read_language_configs_from_excel(&mut workbook)
        .map_err(|e| format!("读取语言配置失败: {}", e))?;
    let sheet_configs =
        read_sheet_configs_from_excel(&app, &mut workbook).map_err(|e| e.to_string())?;

    send_progress(
        &app,
        &format!(
            "读取到 {} 个语言, {} 个工作表",
            lang_configs.len(),
            sheet_configs.len()
        ),
        LogType::Info,
    )?;

    // 创建导出目录
    let parent = file_path.parent().unwrap_or_else(|| Path::new("."));
    let stem = file_path
        .file_stem()
        .map(|s| s.to_string_lossy())
        .unwrap_or_else(|| "export".into());
    let time_str = Local::now().format("%Y%m%d_%H%M%S").to_string();
    let export_folder_name = format!("{}_{}", stem, time_str);
    let output_dir = parent.join(&export_folder_name);
    fs::create_dir_all(&output_dir).map_err(|e| format!("创建目录失败: {}", e))?;

    send_progress(
        &app,
        &format!("导出文件夹创建完成: {}", output_dir.display()),
        LogType::Success,
    )?;

    let mut all_jsons = vec![];

    for lang_config in &lang_configs {
        send_progress(
            &app,
            &format!("正在处理语言: {}", lang_config.code),
            LogType::Info,
        )?;

        let mut sheet_data_map: IndexMap<String, IndexMap<String, String>> = IndexMap::new();

        for sheet_config in &sheet_configs {
            let range = match workbook.worksheet_range(&sheet_config.name) {
                Some(Ok(r)) => r,
                Some(Err(e)) => {
                    send_progress(
                        &app,
                        &format!("⚠️ 读取工作表 {} 失败: {}", sheet_config.name, e),
                        LogType::Warning,
                    )?;
                    continue;
                }
                None => {
                    send_progress(
                        &app,
                        &format!("⚠️ 找不到工作表: {}", sheet_config.name),
                        LogType::Warning,
                    )?;
                    continue;
                }
            };

            let header_row = match range.rows().next() {
                Some(h) => h,
                None => continue,
            };

            let lang_col = header_row
                .iter()
                .position(|c| get_cell_string(c) == lang_config.code);
            let lang_col = match lang_col {
                Some(c) => c,
                None => continue,
            };

            let mut temp: IndexMap<String, String> = IndexMap::new();

            for (row_idx, row) in range.rows().enumerate().skip(1) {
                let key = get_cell_string(&row[0]);
                if key.is_empty() {
                    continue;
                }

                let value = if row.len() > lang_col {
                    get_cell_string(&row[lang_col])
                } else {
                    String::new()
                };

                if value.is_empty() {
                    send_progress(
                        &app,
                        &format!(
                            "空值警告 Sheet: '{}' 行: {} 列: '{}' Key: '{}'",
                            sheet_config.name,
                            row_idx + 1,
                            lang_config.code,
                            key
                        ),
                        LogType::Warning,
                    )?;
                }

                if let Err(err) = check_placeholders(&value) {
                    let msg = format!(
                        "占位符校验失败 Sheet: '{}' 行: {} Key: '{}' 值: '{}' 错误: {}",
                        sheet_config.name,
                        row_idx + 1,
                        key,
                        value,
                        err
                    );
                    let _ = fs::remove_dir_all(&output_dir);
                    return Err(msg);
                }

                temp.insert(key, value);
            }

            sheet_data_map.insert(sheet_config.name.clone(), temp);
        }

        // 合并 sheet 数据
        let mut final_json = JsonValue::new_object();
        for sheet_config in &sheet_configs {
            if let Some(temp) = sheet_data_map.get(&sheet_config.name) {
                if sheet_config.sheet_type.as_deref() == Some("root") {
                    for (k, v) in temp {
                        final_json[k] = v.clone().into();
                    }
                } else {
                    let mut sheet_obj = JsonValue::new_object();
                    for (k, v) in temp {
                        sheet_obj[k] = v.clone().into();
                    }
                    final_json[sheet_config.name.clone()] = sheet_obj;
                }
            }
        }

        // 写入文件
        let json_str = final_json.pretty(2);
        let output_path = output_dir.join(format!("{}.json", lang_config.code));
        fs::write(&output_path, json_str).map_err(|e| e.to_string())?;
        send_progress(
            &app,
            &format!("✅ 已导出语言文件: {}", output_path.display()),
            LogType::Success,
        )?;
        all_jsons.push(output_path);
    }

    // 压缩导出文件夹
    let zip_path = output_dir.with_extension("zip");
    send_progress(&app, "正在压缩导出文件夹...", LogType::Info)?;
    zip_directory(&output_dir, &zip_path)?;
    send_progress(
        &app,
        &format!("✅ 已压缩文件夹为: {}", zip_path.display()),
        LogType::Success,
    )?;

    // 删除原始文件夹
    if let Err(e) = fs::remove_dir_all(&output_dir) {
        send_progress(&app, &format!("⚠️ 删除文件夹失败: {}", e), LogType::Warning)?;
    }

    Ok(format!(
        "完成导出 {} 个语言文件并已压缩为 {:?}",
        all_jsons.len(),
        zip_path
    ))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![convert_excel_to_json])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
