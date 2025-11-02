#!/bin/bash

set -e

echo "🚀 开始构建 Tauri 应用..."

# 确保 tauri-cli 安装
if ! command -v cargo-tauri &> /dev/null; then
  echo "⚙️ 未检测到 cargo-tauri，正在安装..."
  cargo install tauri-cli
fi

# 检查 tauri 版本
cargo tauri --version

echo ""
echo "🧩 构建 macOS DMG..."
cargo tauri build

echo ""
echo "🪟 构建 Windows EXE..."
# 如果还没安装 cross 工具链
rustup target add x86_64-pc-windows-gnu || true

# 构建 Windows 版本
cargo tauri build --target x86_64-pc-windows-gnu

echo ""
echo "✅ 构建完成！"

echo ""
echo "📦 打包结果位置："
echo "  - macOS DMG: src-tauri/target/release/bundle/dmg/"
echo "  - Windows EXE (MSI): src-tauri/target/x86_64-pc-windows-gnu/release/bundle/msi/"

echo ""
echo "🎉 所有构建任务已完成！"
