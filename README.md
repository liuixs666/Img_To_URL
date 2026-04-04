# 图片转CDN链接生成器

这是一个自动化工具，用于扫描图片文件夹，生成jsDelivr CDN链接，并自动提交到GitHub仓库。

## 功能特点

1. **自动扫描**：扫描`images/`文件夹下的所有图片文件
2. **CDN链接生成**：根据图片相对路径生成jsDelivr CDN链接
3. **Excel导出**：在`output/`文件夹生成包含文件名和CDN链接的Excel文件
4. **Git自动化**：自动执行`git add .`、`git commit -m 'Update'`和`git push origin main`

## 使用方法

### 1. 前置准备

#### 1.1 安装依赖
```bash
pip install -r requirements.txt
```

#### 1.2 设置GitHub仓库
1. 在GitHub上创建一个新的仓库
2. 设置git远程地址：
```bash
git remote add origin https://github.com/你的用户名/你的仓库名.git
```

#### 1.3 确保分支为main
```bash
git branch -M main
```

### 2. 添加图片
将需要处理的图片放入`images/`文件夹中。

支持的文件格式：`.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.webp`, `.svg`

### 3. 运行脚本
```bash
cd tools
python core_engine.py
```

## 文件结构
```
Img_To_URL/
├── images/          # 存放图片文件
├── output/          # 生成的Excel文件
├── tools/           # 工具脚本
│   └── core_engine.py
├── README.md        # 说明文档
├── requirements.txt # 依赖列表
└── .git/            # Git仓库
```

## CDN链接格式
生成的CDN链接格式为：
```
https://cdn.jsdelivr.net/gh/你的用户名/你的仓库名/路径/文件名
```

例如：
```
https://cdn.jsdelivr.net/gh/username/repository/images/photo.jpg
```

## 注意事项

1. **首次运行前**：确保已设置git远程(`git remote add origin`)
2. **Git权限**：确保有权限推送到仓库
3. **网络连接**：需要网络连接来推送代码到GitHub
4. **文件路径**：图片路径中不要包含中文字符或特殊字符
5. **Excel文件**：每次运行会在`output/`文件夹生成带时间戳的新Excel文件

## 故障排除

### 错误：未找到git远程 'origin'
```bash
git remote add origin https://github.com/你的用户名/你的仓库名.git
```

### 错误：需要pandas库
```bash
pip install pandas openpyxl
```

### 错误：推送权限不足
检查GitHub访问令牌或SSH密钥设置

## 更新日志
- 2026-04-04: 初始版本发布