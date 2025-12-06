#!/bin/bash
# Intel Mac构建脚本

set -e

echo "开始构建Intel版本..."

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3"
    exit 1
fi

# 检查PyInstaller
if ! python3 -c "import PyInstaller" 2>/dev/null; then
    echo "安装PyInstaller..."
    pip3 install pyinstaller
fi

# 检查依赖
echo "检查依赖..."
pip3 install requests

# 切换到项目根目录
cd "$(dirname "$0")/.."

# 创建虚拟环境（可选）
if [ ! -d "venv_intel" ]; then
    echo "创建Intel虚拟环境..."
    python3 -m venv venv_intel
fi

# 激活虚拟环境
source venv_intel/bin/activate

# 安装依赖
pip install -r requirements.txt 2>/dev/null || pip install requests

# 清理之前的构建
echo "清理之前的构建文件..."
rm -rf build/build
rm -rf build/dist
rm -f build/etf_api_caller.spec

# 执行构建
echo "开始PyInstaller构建..."
cd build
pyinstaller --clean etf_caller.spec --target-arch=x86_64

# 检查构建结果
if [ -f "dist/etf_api_caller" ]; then
    echo "✅ Intel版本构建成功!"
    echo "可执行文件: $(pwd)/dist/etf_api_caller"
    
    # 测试构建结果
    echo "测试可执行文件..."
    ./dist/etf_api_caller --test --codes "159915" || echo "⚠️  测试失败，但构建完成"
    
    # 复制到项目根目录
    cp dist/etf_api_caller ../etf_api_caller_intel
    echo "已复制到: ../etf_api_caller_intel"
else
    echo "❌ 构建失败"
    exit 1
fi

# 显示文件信息
echo "文件信息:"
ls -la dist/etf_api_caller
file dist/etf_api_caller

deactivate
echo "构建完成!"
