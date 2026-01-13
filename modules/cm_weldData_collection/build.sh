#!/bin/bash

echo "========================================"
echo "正在打包 weldData_collection 为可执行文件"
echo "========================================"
echo ""

cd weldData_collection

echo "正在发布自包含单文件可执行程序..."
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

if [ $? -eq 0 ]; then
    echo ""
    echo "========================================"
    echo "打包成功！"
    echo "========================================"
    echo "可执行文件位置："
    echo "bin/Release/net8.0/win-x64/publish/weldData_collection.exe"
    echo ""
    echo "该文件可以在任何 Windows x64 系统上运行，无需安装 .NET 运行时"
else
    echo ""
    echo "========================================"
    echo "打包失败，请检查错误信息"
    echo "========================================"
fi

