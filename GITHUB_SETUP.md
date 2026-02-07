# GitHub 自动发布配置说明

## 配置完成

已为你创建 GitHub Actions 自动发布工作流，仓库地址：
**https://github.com/chxwy/Y2Tool**

## 文件说明

| 文件 | 说明 |
|------|------|
| `.github/workflows/release.yml` | GitHub Actions 工作流配置 |
| `.github/release_template.md` | 发布说明模板 |
| `docs/version.json` | 版本信息文件（自动更新） |

## 发布新版本

### 方法1：推送标签（推荐）

```bash
# 1. 确保代码已提交
git add .
git commit -m "准备发布 v1.9.1"

# 2. 创建标签
git tag v1.9.1

# 3. 推送标签到 GitHub
git push origin v1.9.1
```

推送后，GitHub Actions 会自动：
1. 构建 Windows 可执行文件
2. 创建 ZIP 压缩包
3. 计算 SHA256 哈希
4. 生成 version.json
5. 创建 GitHub Release
6. 上传所有文件

### 方法2：手动触发

1. 打开仓库页面：https://github.com/chxwy/Y2Tool
2. 点击 "Actions" 标签
3. 选择 "Build and Release" 工作流
4. 点击 "Run workflow"
5. 输入版本号（如 `1.9.1`）
6. 点击 "Run workflow"

## 更新流程

### 首次发布（v1.9.0）

由于自动更新功能需要 `version.json` 存在，首次发布需要手动操作：

```bash
# 1. 本地构建测试
python build_and_release.py

# 2. 上传文件到 GitHub Release（手动创建 Release）
# - 访问 https://github.com/chxwy/Y2Tool/releases
# - 点击 "Create a new release"
# - 上传 release/ 目录中的文件

# 3. 提交 version.json 到仓库
git add docs/version.json
git commit -m "添加版本信息"
git push
```

### 后续版本（v1.9.1+）

只需推送标签，全部自动化：

```bash
git tag v1.9.1
git push origin v1.9.1
```

## 更新日志配置

发布新版本时，编辑 `.github/release_template.md` 填写更新内容，或使用 GitHub 的自动生成发布说明功能。

## 自动更新工作原理

```
用户启动程序
    │
    ▼
程序访问 https://raw.githubusercontent.com/chxwy/Y2Tool/main/docs/version.json
    │
    ▼
对比本地版本 vs 远程版本
    │
    ├── 有更新 → 显示更新对话框 → 用户确认 → 下载 → 安装 → 重启
    │
    └── 无更新 → 静默继续
```

## 备用更新源（可选）

如果 GitHub 访问不稳定，可以配置 Gitee 镜像：

1. 在 Gitee 导入 GitHub 仓库
2. 修改 `update_module.py` 中的 `BACKUP_CHECK_URL`
3. 同步发布到 Gitee Releases

## 常见问题

### Q: GitHub Actions 构建失败？

A: 检查以下几点：
- `Y2订单处理辅助工具1.9.spec` 文件是否存在
- 依赖包是否完整
- 查看 Actions 日志定位错误

### Q: 用户无法检查到更新？

A: 检查：
- `docs/version.json` 是否已推送到 main 分支
- 版本号格式是否正确（如 `1.9.0` 不带 v）
- URL 是否可访问（尝试浏览器直接访问）

### Q: 如何测试更新功能？

A: 本地测试方法：
```python
# 修改 CURRENT_VERSION 为旧版本号
CURRENT_VERSION = "1.8.0"

# 运行程序，应该能检测到更新
```

## 下一步操作

1. **初始化仓库**
   ```bash
   git init
   git remote add origin https://github.com/chxwy/Y2Tool.git
   git add .
   git commit -m "初始提交"
   git push -u origin main
   ```

2. **首次发布 v1.9.0**
   - 本地运行 `python build_and_release.py`
   - 手动创建 GitHub Release 并上传文件
   - 提交 `docs/version.json`

3. **后续版本**
   - 修改代码
   - 推送标签 `git tag v1.9.1 && git push origin v1.9.1`
   - 等待自动构建完成
