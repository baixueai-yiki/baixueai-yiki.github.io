# baixueai-yiki.github.io
萝卜网
baixueai-yiki.github.io/
├── .git/
├── Srticles/   # 子页
│   ├── About.html
│   ├── Demo.html
│   ├── Share.html
│   └── Recommend.html
├── Content/    # 资源
│   ├── head/
│   ├── Icon/
│   └── Images/
├── css/        # 样式文件
│   ├── animation.css   # 动画
│   ├── component.css   # 组件
│   ├── layout.css      # 页面骨架
│   └── tool.css        # 小工具
├── Docs/       # 文档
├── Share/      # 分享页供下载的文件
│   ├── About.html
│   ├── Demo.html
│   ├── Share.html
│   └── Recommend.html
├── index.html  # 页面结构（HTML）
├── main.js     # 页面逻辑
└── style.css   # 网页美化
## 怎么加一篇新文章

1. 在 `Content/Posts/<抽屉>/<分组>/` 里放一个 **PDF**，文件夹名就是网页上的分组名。
2. 传上去就行。GitHub Actions 会刷新 `index.json` 再提交回来；等一两分钟网页就有了。
3. 按钮上的字 = PDF 的文件名，点开是一个只用来显示这份 PDF 的页面。

Word 仍然照常写，传之前另存成 PDF 就行（Word 里「文件 → 导出 → PDF」一步）。这样做的好处是网页上看到的和你写的完全一致，不用经过任何转换。

抽屉对应四个 tab：`share` 分享、`recommend` 安利。目前只有安利接了这个流程。

想在本机先试也可以（可选）：

```
python tools/build_posts.py
```
