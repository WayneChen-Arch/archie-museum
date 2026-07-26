# Archie Museum

A minimalist online art gallery for Archie's drawings.

## 作品资料 Excel

仓库根目录的 [`works.xlsx`](./works.xlsx) 汇总全部作品的名称、年份、创作媒介与一句话介绍（简体中文为主）。

1. 下载并编辑黄色列：`名称` / `年份` / `创作媒介` / `一句话介绍`（以及可选的创作过程说明）。
2. 勿改 `序号`、`图片文件`。
3. 英文列可按需修改；繁体无需手填。
4. 上传覆盖 GitHub 上的 `works.xlsx` 后，运行：

```bash
python3 scripts/sync-works-from-excel.py import
```

同步规则：简体写入网页简体；繁体由简体自动转换；英文优先用 Excel 英文列，留空则保留原英文。

从网页重新导出 Excel：

```bash
python3 scripts/sync-works-from-excel.py export
```

## Deploy to GitHub Pages

1. Open PowerShell in this folder.
2. Run:

```powershell
.\deploy.ps1
```

3. Sign in to GitHub when prompted.
4. Open the printed URL on your phone after 1-3 minutes.

If the repository already exists, push manually:

```powershell
git remote add origin https://github.com/<your-username>/archie-museum.git
git push -u origin master
```

Then enable **GitHub Pages** in repository settings: branch `master`, folder `/ (root)`.
