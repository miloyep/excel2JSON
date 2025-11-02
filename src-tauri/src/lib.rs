use calamine::{open_workbook, Reader, Xlsx};
use chrono::Local;
use serde::Serialize;
use serde_json::json;
use std::fs;
use tauri::{AppHandle, Emitter};

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

#[tauri::command]
async fn convert_excel_to_json(app: AppHandle, path: String) -> Result<String, String> {
    send_progress(&app, &format!("开始处理文件: {}", path), LogType::Info)?;

    let file_path = std::path::PathBuf::from(&path);
    let mut workbook: Xlsx<_> =
        open_workbook(&file_path).map_err(|e| format!("打开文件失败: {}", e))?;

    let sheet_lang_name = "导出语言管理";
    let sheet_obj_name = "导出sheet管理";

    let data_lang = workbook
        .worksheet_range(sheet_lang_name)
        .ok_or("找不到 Sheet: 导出语言管理")?
        .map_err(|_| "读取 sheet_lang 失败")?;

    let data_obj = workbook
        .worksheet_range(sheet_obj_name)
        .ok_or("找不到 Sheet: 导出sheet管理")?
        .map_err(|_| "读取 sheet_obj 失败")?;

    let parent = file_path.parent().unwrap_or(std::path::Path::new("."));
    let folder_name = format!(
        "{}_{}",
        file_path.file_stem().unwrap().to_string_lossy(),
        Local::now().format("%Y%m%d_%H%M%S")
    );
    let output_dir = parent.join(&folder_name);

    fs::create_dir_all(&output_dir).map_err(|e| e.to_string())?;

    let mut all_jsons = vec![];

    for row_lang in data_lang.rows() {
        let lang = row_lang[0].to_string();
        let mut obj = serde_json::Map::new();

        send_progress(&app, &format!("正在处理语言: {}", lang), LogType::Info)?;

        for row_obj in data_obj.rows() {
            let sheet_name = row_obj[0].to_string();
            let sheet_type = row_obj[1].to_string();

            let range = workbook
                .worksheet_range(&sheet_name)
                .ok_or(format!("找不到 Sheet: {}", sheet_name))?
                .map_err(|_| "读取 sheet 失败")?;

            // 找到对应语言列
            let mut col_of_lang = None;
            if let Some(header_row) = range.rows().next() {
                for (i, cell) in header_row.iter().enumerate() {
                    if cell.to_string() == lang {
                        col_of_lang = Some(i);
                    }
                }
            }
            let col_of_lang = match col_of_lang {
                Some(c) => c,
                None => continue,
            };

            let mut temp = serde_json::Map::new();

            for (row_idx, row) in range.rows().skip(1).enumerate() {
                let key = row[0].to_string();
                let value = row[col_of_lang].to_string();

                // 空值检查
                if value.trim().is_empty() {
                    let _ = send_progress(
                        &app,
                        &format!(
                            "空值警告 Sheet: '{}' 行: {} 列: '{}' Key: '{}'",
                            sheet_name,
                            row_idx + 2,
                            lang,
                            key
                        ),
                        LogType::Warning,
                    );
                }

                // 占位符校验
                if let Err(err) = check_placeholders(&value) {
                    let _ = fs::remove_dir_all(&output_dir);
                    return Err(format!(
                        "占位符校验失败 Sheet: '{}' 行: {} 列: '{}' Key: '{}' 值: '{}', 错误: {}",
                        sheet_name,
                        row_idx + 2,
                        lang,
                        key,
                        value,
                        err
                    ));
                }

                if sheet_type == "root" {
                    obj.insert(key, json!(value));
                } else {
                    temp.insert(key, json!(value));
                }
            }

            if sheet_type != "root" {
                obj.insert(sheet_name, json!(temp));
            }
        }

        let json_str = serde_json::to_string_pretty(&obj).unwrap();
        let output_path = output_dir.join(format!("{}.json", lang));
        fs::write(&output_path, &json_str).map_err(|e| e.to_string())?;

        all_jsons.push(output_path.clone());

        send_progress(
            &app,
            &format!("已导出语言文件: {:?}", output_path),
            LogType::Success,
        )?;
    }

    send_progress(
        &app,
        &format!("完成导出 {} 个语言文件到 {:?}", all_jsons.len(), output_dir),
        LogType::Success,
    )?;

    Ok(format!(
        "已导出 {} 个语言文件到 {:?}",
        all_jsons.len(),
        output_dir
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
