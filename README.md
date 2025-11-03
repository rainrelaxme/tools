Tools
===============
当前最新版本： 0.9

[![AUR](https://img.shields.io/badge/license-Apache%20License%202.0-blue.svg)](https://github.com/zhangdaiscott/jeecg-boot/blob/master/LICENSE)


## 简介
一些小工具

## 安装与使用

 > 环境要求: 版本要求python 3.12.11。

- Get the project code

```bash
git clone https://github.com/rainrelaxme/tools.git
```

- Installation dependencies

```bash
cd tools

conda activate General  # conda环境

pip install -r requirements.txt # 安装包
```
-  Create package

```bash
cd tools\scripts

pyinstaller package.spec
```

## 功能清单
 > * Excel压缩体积：压缩其中的图片
 > * 


## 项目结构

```
├─config/
├─data/
│  ├─input/
│  ├─output/
│  ├─sample/
│  ├─temp/
│  ├─template/
|  └─test/
└─docs/
└─logs/
├─scripts/
├─src/
│  ├─common/
│  ├─component/
│  ├─service/
│  ├─test/
|  ├─view/
|  |  ├─demo/
|  |  ├─production/
|  |  └─side_project/
|  └─run.py
├─tests/
├─.gitignore
├─main.py
├─pyproject.toml
├─README.md
└─requirements.txt
```


