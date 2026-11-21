"use strict";

// 下雪效果
function createSnow()
{
    const snow = document.createElement("div");
    snow.classList.add("snowflake");
    snow.innerText = "❄";
    snow.style.left = Math.random() * 100 + "vw";
    snow.style.fontSize = Math.random() * 12 + 8 + "px";
    snow.style.opacity = Math.random();

    const duration = Math.random() * 6 + 6;
    snow.style.animationDuration = duration + "s";

    document.body.appendChild(snow);

    setTimeout(() =>
    {
        snow.remove();
    }, duration * 1000);
}
setInterval(createSnow, 150);

// 随机顶部图片
const headerImages = [
  "Content/Images/header-image-1.png","Content/Images/header-image-2.png",
  "Content/Images/header-image-3.png","Content/Images/header-image-4.png",
  "Content/Images/header-image-5.png","Content/Images/header-image-6.png",
  "Content/Images/header-image-7.png","Content/Images/header-image-8.png",
  "Content/Images/header-image-9.png","Content/Images/header-image-10.png"
];

const header = document.querySelector(".top-banner");

if (header)
{
    const randomImage = headerImages[Math.floor(Math.random() * headerImages.length)];
    header.style.backgroundImage = `url('${randomImage}')`;
}

// Tab 页面切换
const buttons = document.querySelectorAll(".tab-btn");
const contents = document.querySelectorAll(".tab-content");
const loadedTabs = {};
const tabFiles = {
    tab1: "articles/About.html",
    tab2: "articles/Demo.html",
    tab3: "articles/Share.html",
    tab4: "articles/Recommend.html"
};

function openTab(tabId)
{
    contents.forEach(content =>
    {
        content.classList.remove("active");
    });

    buttons.forEach(button =>
    {
        button.classList.remove("active");
    });
    document.getElementById(tabId).classList.add("active");

    document
        .querySelector(`[data-tab="${tabId}"]`)
        .classList.add("active");

    if (!loadedTabs[tabId])
    {
        loadTabHTML(tabId).then(() =>
        {
            const content = document.getElementById(tabId);
            bindArticles(content);
            return renderPostList(content, tabId);
        }).then(() =>
        {
            generateSidebarTOC(tabId);
        });
        loadedTabs[tabId] = true;
    }
    else
    {
        generateSidebarTOC(tabId);
    }
    location.hash = tabId;
}

buttons.forEach(button =>
{
    button.addEventListener("click", () =>
    {
        openTab(button.dataset.tab);
    });
});

function loadTabHTML(tabId)
{
    return fetch(tabFiles[tabId])
    .then(response => response.text())
    .then(html =>
    {
        document.getElementById(tabId).innerHTML = html;
    })
    .catch(() =>
    {
        document.getElementById(tabId).innerHTML =
            "<p>加载失败……</p>";
    });
}

// 收起的文章也要知道自己在不在加载中
const articleState = new WeakMap();

function articleBody(details)
{
    let body = details.querySelector(".article-body");

    if (!body)
    {
        body = document.createElement("div");
        body.className = "article-body";
        details.insertBefore(
            body,
            details.querySelector("summary").nextSibling
        );
    }
    return body;
}

// 展开才去读文章文件，读过就不再读
function loadArticle(details)
{
    const file = details.dataset.file;

    if (!file)
    {
        return Promise.resolve();
    }

    const state = articleState.get(details);

    if (state === "loaded")
    {
        return Promise.resolve();
    }
    if (state === "loading")
    {
        return details.loadPromise;
    }

    const body = articleBody(details);

    body.classList.remove("is-failed");
    body.classList.add("is-loading");
    body.textContent = "正在读取……";
    articleState.set(details, "loading");

    const task = fetch(file)
    .then(response =>
    {
        if (!response.ok)
        {
            throw new Error(response.status + " " + file);
        }
        return response.text();
    })
    .then(html =>
    {
        body.classList.remove("is-loading");
        body.innerHTML = html;
        articleState.set(details, "loaded");
    })
    .catch(error =>
    {
        body.classList.remove("is-loading");
        body.classList.add("is-failed");
        body.textContent = `这篇没读出来：${file}`;
        articleState.set(details, "failed");
        console.error(error);
    });

    details.loadPromise = task;
    return task;
}

// 给骨架里带 data-file 的折叠块挂上展开事件
function bindArticles(root)
{
    root.querySelectorAll("details[data-file]").forEach(details =>
    {
        if (details.dataset.bound)
        {
            return;
        }
        details.dataset.bound = "1";

        details.addEventListener("toggle", () =>
        {
            if (details.open)
            {
                loadArticle(details);
            }
        });
    });
}

// 标签页里的文章列表：从 Content/Posts/<tab>/index.json 画分组和按钮
function renderPostList(root, tabId)
{
    const container = root.querySelector("[data-post-list]");

    if (!container)
    {
        return Promise.resolve();
    }

    const listName = container.dataset.postList || tabId;
    const url = `Content/Posts/${listName}/index.json`;

    return fetch(url, { cache: "no-store" })
    .then(response =>
    {
        if (!response.ok)
        {
            throw new Error(response.status + " " + url);
        }
        return response.json();
    })
    .then(data =>
    {
        container.innerHTML = "";

        if (!data.groups || !data.groups.length)
        {
            container.innerHTML =
                '<p class="post-loading">这个抽屉还空着，把 word 丢进 Content/Posts/' +
                listName + "/<分组>/ 就会出现在这里</p>";
            return;
        }

        data.groups.forEach(group =>
        {
            const details = document.createElement("details");
            details.className = "post-group";
            details.dataset.tocType = "parent";
            details.dataset.toc = group.title;
            details.open = true;

            const summary = document.createElement("summary");
            summary.innerHTML = `<strong>${group.title}</strong>`;

            const list = document.createElement("div");
            list.className = "post-buttons";

            group.items.forEach(item =>
            {
                const link = document.createElement("a");
                link.className = "post-link";
                link.href = `article.html?f=${encodeURIComponent(item.path)}&t=${tabId}`;
                link.textContent = item.title;
                list.appendChild(link);
            });

            details.appendChild(summary);
            details.appendChild(list);
            container.appendChild(details);
        });
    })
    .catch(error =>
    {
        container.innerHTML =
            '<p class="post-loading">文章列表没读出来，可能是还没生成</p>';
        console.error(error);
    });
}

// 自动生成目录
function generateSidebarTOC(tabId)
{
    const toc = document.getElementById("toc");
    toc.querySelectorAll("li:not(:first-child)")
        .forEach(item => item.remove());
    const content = document.getElementById(tabId);
    setTimeout(() =>
    {
        const parents = content.querySelectorAll(
            "details[data-toc-type='parent']"
        );


        parents.forEach(parent => {

            const li = document.createElement("li");
            const link = document.createElement("a");

            link.href = "javascript:void(0)";
            link.textContent = parent.dataset.toc;

            li.appendChild(link);


            const children = parent.querySelectorAll(
                "details[data-toc-type='child'], a.post-link"
            );


            if (children.length) 
            {
                const ul = document.createElement("ul");


                children.forEach(child =>
                {

                    const childLi = document.createElement("li");
                    const childLink = document.createElement("a");
                    const isPostLink = child.tagName === "A";

                    childLink.href = "javascript:void(0)";
                    childLink.textContent = isPostLink
                        ? child.textContent
                        : child.dataset.toc;

                    childLink.onclick = isPostLink
                        ? () =>
                        {
                            location.href = child.getAttribute("href");
                        }
                        : () =>
                        {
                            if (!child.open)
                            {
                                child.open = true;
                            }
                            // 先等内容到位，再滚过去，否则会停在半路
                            loadArticle(child).then(() =>
                            {
                                child.scrollIntoView(
                                {
                                    behavior: "smooth",
                                    block: "start"
                                });
                            });
                        };
                    childLi.appendChild(childLink);
                    ul.appendChild(childLi);
                });
                li.appendChild(ul);
                link.onclick = () => {
                    ul.classList.toggle("expanded");
                };
            }
            toc.appendChild(li);
        });
    }, 100);
}

// 初始页面
const hash = location.hash.replace("#", "");
if (hash && document.getElementById(hash)) 
{
    openTab(hash);
}
else
{
    openTab("tab1");
}

// 返回顶部
document
.getElementById("back-to-top")
.onclick = () =>
{
    window.scrollTo(
    {
        top: 0,
        behavior: "smooth"
    });
};

// 侧边目录：鼠标靠近展开、离开两秒后收起；点把手常驻展开，再点取消
const sidebar = document.querySelector(".sidebar");
const sidebarToggle = document.getElementById("sidebar-toggle");
const SIDEBAR_HIDE_DELAY = 500;

let pinned = false;
let hideTimer = null;
let hovering = false;

function showSidebar()
{
    sidebar.classList.add("show");
}

function hideSidebarSoon()
{
    if (pinned || hovering)
    {
        return;
    }
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() =>
    {
        hideTimer = null;
        if (pinned || hovering)
        {
            return;
        }
        sidebar.classList.remove("show");
    }, SIDEBAR_HIDE_DELAY);
}

function setPinned(value)
{
    pinned = value;
    sidebar.classList.toggle("pinned", pinned);
    sidebarToggle.textContent = pinned ? "⮜" : "⮞";
    sidebarToggle.title = pinned ? "点击取消常驻" : "点击常驻展开";

    if (pinned)
    {
        clearTimeout(hideTimer);
        hideTimer = null;
        showSidebar();
    }
    else
    {
        hideSidebarSoon();
    }
}

sidebarToggle.addEventListener("click", () =>
{
    setPinned(!pinned);
});

sidebar.addEventListener("pointerenter", event =>
{
    if (event.pointerType !== "mouse")
    {
        return;
    }
    hovering = true;
    clearTimeout(hideTimer);
    hideTimer = null;
    showSidebar();
});

sidebar.addEventListener("pointerleave", event =>
{
    if (event.pointerType !== "mouse")
    {
        return;
    }
    hovering = false;
    hideSidebarSoon();
});

document.addEventListener("mousemove", event =>
{
    if (pinned)
    {
        return;
    }

    const nearRight =
        event.clientX > window.innerWidth - 30;

    if (nearRight)
    {
        clearTimeout(hideTimer);
        hideTimer = null;
        showSidebar();
    }
    else if (!hovering && sidebar.classList.contains("show") && !hideTimer)
    {
        hideSidebarSoon();
    }
});

// 桌宠
const pet = document.getElementById("desktop-pet");
const idleGif = "Content/Images/cat1.gif";
const animGif = "Content/Images/cat.gif";
pet.style.backgroundImage = `url('${idleGif}')`;
pet.onclick = () =>
{
    pet.style.backgroundImage = `url('${animGif}')`;
    setTimeout(() =>
    {
        pet.style.backgroundImage = `url('${idleGif}')`;
    }, 2000);
};  