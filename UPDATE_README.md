# Y2订单处理辅助工具 - 自动更新系统

## 概述

本工具已实现远程自动更新功能，用户可以通过以下方式获取最新版本：

1. **自动检查** - 程序启动后3秒自动检查更新（静默模式）
2. **手动检查** - 在设置页面的"作者信息"标签页中点击"检查更新"按钮

## 更新流程

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  检查版本    │────►│  下载更新包  │────►│  安装更新   │
│  (version)  │     │  (download) │     │ (updater)  │
└─────────────┘     └─────────────┘     └─────────────┘
```

## 服务端配置

### 1. 托管 version.json

将 `version.json` 文件托管到可访问的 URL（如 GitHub、Gitee、自己的服务器）：

```json
{
  "version": "1.9.0",
  "min_version": "1.8.0",
  "download_url": "https://your-domain.com/releases/Y2订单处理辅助工具1.9.zip",
  "changelog": [
    "新增远程自动更新功能",
    "优化图片处理性能"
  ],
  "force_update": false,
  "file_size": 15234567,
  "hash": "sha256:abc123..."
}
```

### 2. 配置更新源

修改 `update_module.py` 中的 URL：

```python
# 主更新源
VERSION_CHECK_URL = "https://your-domain.com/version.json"

# 备用更新源（可选）
BACKUP_CHECK_URL = "https://gitee.com/yourname/Y2Tool/raw/main/version.json"
```

## 发布新版本

### 方法1：使用构建脚本（推荐）

```bash
# 完整构建并生成发布文件
python build_and_release.py

# 仅清理构建目录
python build_and_release.py --clean
```

### 方法2：手动构建

```bash
# 1. 构建应用
pyinstaller Y2订单处理辅助工具1.9.spec --clean

# 2. 打包发布文件
cd dist
zip -r Y2订单处理辅助工具1.9.zip Y2订单处理辅助工具1.9

# 3. 计算哈希值
sha256sum Y2订单处理辅助工具1.9.zip

# 4. 更新 version.json 并上传
```

## 版本号规则

采用语义化版本控制（SemVer）：

- **主版本号** - 重大功能变更（如 1.x → 2.0）
- **次版本号** - 新功能添加（如 1.8 → 1.9）
- **修订号** - Bug修复（如 1.9.0 → 1.9.1）

## 文件说明

| 文件 | 说明 |
|------|------|
| `update_module.py` | 更新检查模块（版本检查、下载、UI） |
| `updater.py` | 更新助手程序（文件替换、重启） |
| `version.json` | 版本信息配置文件 |
| `build_and_release.py` | 自动构建发布脚本 |

## 注意事项

1. **网络要求** - 更新功能需要网络连接
2. **权限要求** - Windows 可能需要管理员权限才能替换程序文件
3. **备份机制** - 更新前会自动备份旧版本到临时目录
4. **回滚方案** - 如果更新失败，可以从备份目录恢复

## 故障排除

### 检查更新失败

- 检查网络连接
- 确认 version.json URL 可访问
- 查看防火墙设置

### 下载失败

- 检查 download_url 是否正确
- 确认文件托管服务器可用
- 检查磁盘空间

### 安装失败

- 确保程序有写入权限
- 关闭杀毒软件重试
- 手动下载更新包解压覆盖

## 自定义更新服务器

如果不想使用 GitHub/Gitee，可以自建更新服务器：

1. **静态文件服务器** - Nginx/Apache 托管 version.json 和更新包
2. **对象存储** - 阿里云 OSS、腾讯云 COS、AWS S3 等
3. **CDN 加速** - 使用 CDN 分发更新包，提高下载速度

## 安全建议

1. **HTTPS 优先** - 使用 HTTPS 防止中间人攻击
2. **文件签名** - 使用 SHA256 校验文件完整性
3. **版本白名单** - 设置 min_version 防止降级攻击
4. **强制更新** - 重要安全更新可设置 force_update: true
