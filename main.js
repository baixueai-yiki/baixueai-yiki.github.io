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
                "details[data-toc-type='child']"
            );


            if (children.length) 
            {
                const ul = document.createElement("ul");


                children.forEach(child =>
                {

                    const childLi = document.createElement("li");
                    const childLink = document.createElement("a");

                    childLink.href = "javascript:void(0)";
                    childLink.textContent = child.dataset.toc;

                    childLink.onclick = () =>
                    {

                        if (!child.open)
                        {
                            child.open = true;
                        }
                        child.scrollIntoView(
                        {
                            behavior: "smooth",
                            block: "start"
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

// 侧边目录显示
const sidebar = document.querySelector(".sidebar");
let hideTimer = null;
let hovering = false;

sidebar.addEventListener("mouseenter", () =>
{
    hovering = true;
    sidebar.classList.add("show");
    clearTimeout(hideTimer);
});

sidebar.addEventListener("mouseleave", () =>
{
    hovering = false;
    hideTimer = setTimeout(() => 
    {
        sidebar.classList.remove("show");
    }, 300);
});

document.addEventListener("mousemove", e =>
{
    const nearRight =
        e.clientX > window.innerWidth - 30;
    if (nearRight)
    {
        sidebar.classList.add("show");
        clearTimeout(hideTimer);
    }
    else if (!hovering)
    {
        hideTimer = setTimeout(() =>
        {
            sidebar.classList.remove("show");
        }, 300);
    }
});

// 桌宠
const pet = document.getElementById("desktop-pet");
const idleGif = "Content/images/cat1.gif";
const animGif = "Content/images/cat.gif";
pet.style.backgroundImage = `url('${idleGif}')`;
pet.onclick = () =>
{
    pet.style.backgroundImage = `url('${animGif}')`;
    setTimeout(() =>
    {
        pet.style.backgroundImage = `url('${idleGif}')`;
    }, 2000);
};  