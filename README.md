# Archie Museum

A minimalist online art gallery for Archie's drawings.

## 作品资料 Excel

仓库根目录的 [`works.xlsx`](./works.xlsx) 汇总全部作品的名称、年份、画面尺寸、创作媒介与一句话介绍。
作品图片放在 [`assets/`](./assets/) 目录。

1. 下载并编辑：`名称` / `年份` / `画面尺寸` / `创作媒介` / `一句话介绍` / `创作过程说明`
2. 「图片文件」对应 `assets/` 中不含扩展名的文件名，请勿随意改动
3. 上传覆盖 GitHub 上的 `works.xlsx` 后，运行：

```bash
python3 scripts/sync-works-from-excel.py import
```

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
