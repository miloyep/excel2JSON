import { useEffect, useState, useRef } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import "./App.css";

function App() {
  // 日志数组，每条日志有 message 和 type
  const [logs, setLogs] = useState([]);
  const logRef = useRef(null);

  const appendLog = (msg, type = "info") => {
    setLogs((prev) => [...prev, { message: msg, type }]);
  };

  const handleClear = () => {
    setLogs([]);
  };

  async function handleConvert() {
    const file = await open({
      filters: [{ name: "Excel 文件", extensions: ["xlsx"] }],
    });
    if (!file) return;

    try {
      await invoke("convert_excel_to_json", { path: file });
    } catch (err) {
      appendLog("转换失败：" + err, "error");
    }
  }

  // 自动滚动到底部
  useEffect(() => {
    if (logRef.current) {
      logRef.current.scrollTop = logRef.current.scrollHeight;
    }
  }, [logs]);

  useEffect(() => {
    const unlisten = listen("excel-export-progress", (event) => {
      if (!event.payload) return;

      // event.payload 已经是 { message, type } 对象
      const { message, type } = event.payload;
      appendLog(message, type);
    });

    // 清理监听器
    return () => {
      unlisten.then((f) => f());
    };
  }, []);

  // 日志颜色样式
  const getLogStyle = (type) => {
    switch (type) {
      case "success":
        return { color: "#4caf50" }; // 绿色
      case "warning":
        return { color: "#ff9800" }; // 橙色
      case "error":
        return { color: "#f44336" }; // 红色
      case "info":
      default:
        return { color: "#000000" }; // 黑色
    }
  };

  return (
    <main className="container">
      <h1>📘 Excel 多语言导出工具</h1>
      <div className="row space-x-[20px]">
        <button onClick={handleConvert}>选择Excel并开始导出 JSON</button>
        <button onClick={handleClear}>清空</button>
      </div>
      <div
        ref={logRef}
        style={{
          marginTop: "16px",
          padding: "10px",
          height: "400px",
          overflowY: "auto", // 只允许纵向滚动
          overflowX: "hidden", // 禁止横向滚动
          backgroundColor: "#f0f0f0",
          border: "1px solid #ccc",
          whiteSpace: "pre-wrap", // 保留 \n 换行
          overflowWrap: "break-word", // 超长内容才换行
          wordBreak: "normal", // 避免在符号处自动断行
          textAlign: "left",
          fontFamily: "monospace",
        }}
      >
        {logs.map((log, index) => (
          <div key={index} style={getLogStyle(log.type)}>
            {log.message}
          </div>
        ))}
      </div>
      <p
        style={{
          fontSize: "12px",
          color: "#999",
          fontFamily: "monospace",
        }}
      >
        v1.0.1
      </p>
    </main>
  );
}

export default App;
