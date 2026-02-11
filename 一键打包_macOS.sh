#!/bin/bash
echo "============================================"
echo "   PDF拆分工具 - 一键打包为macOS应用"
echo "============================================"
echo ""

echo "[1/2] 正在安装所需的Python库..."
pip3 install -r requirements.txt
if [ $? -ne 0 ]; then
    echo ""
    echo "安装失败！请确认已安装Python3并且pip3可用。"
    echo "可尝试: brew install python3"
    exit 1
fi

echo ""
echo "[2/2] 正在打包为macOS应用，请稍候..."
pyinstaller --onefile --windowed --name "PDF拆分工具" pdf_splitter.py
if [ $? -ne 0 ]; then
    echo ""
    echo "打包失败！"
    exit 1
fi

echo ""
echo "============================================"
echo "   打包成功！"
echo "   应用位于: dist/PDF拆分工具.app"
echo "   把这个应用拷贝到"应用程序"文件夹即可使用！"
echo "============================================"
echo ""
