# 📐 레이아웃 구조

[🔙 README](../README.md)

## 기본 레이아웃 구조
```html
<!doctype html>
<html lang="ko">
  <head>...</head>
  <body>
    <div id="root">
      <div id="app">
        <!-- Header -->
        <header id="header">
          <div class="header-container">
            <!-- Logo -->
            <div class="header-title">
              <h1 class="logo">...</h1>
            </div>
            <!-- GNB -->
            <div class="header-content">
              <nav id="gnb">...</nav>
            </div>
          </div>
        </header>
        <div id="wrap">
          <div class="scroll-area">
            <!-- Main -->
            <main id="main">
              <div class="main-container">
                <!-- TopBar -->
                <div class="top-bar">...</div>
                <!-- Breadcrumb -->
                <div class="breadcrumb">...</div>
                <!-- Content -->
                <div class="main-content">...</div>
              </div>
            </main>
            <!-- Footer -->
            <footer id="footer">...</footer>
          </div>
        </div>
      </div>
    </div>
  </body>
</html>

```

## head 구조
❗️`<head />` 변경 필요❗
> 퍼블리싱 작업을 효율적으로 진행하기 위해 CSS 및 JavaScript를 import 방식으로 적용했으나, 실제 개발 시에는 외부 파일 적용 방식으로의 변경 필요<br/>
```diff
<head>
    <!-- MetaData -->
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="CHA Biotech BECARE 건진센터 통합관리자" />
    <meta name="keywords" content="CHA, Biotech, 차바이오텍, BECARE, 건진센터, 통합관리자" />
    <title>BECARE 통합관리자</title>
    
    <!-- Fonts -->
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Black.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Bold.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-ExtraBold.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-ExtraLight.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Light.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Medium.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Regular.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-SemiBold.woff2" as="font" type="font/woff2" crossorigin />
    <link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Thin.woff2" as="font" type="font/woff2" crossorigin />
    
    <!-- StyleSheet -->
-   <link rel="stylesheet" href="/assets/styles/index.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/fonts.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/reset.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/vars.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/common.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/layouts.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/pages.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/components/form.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/components/ui.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/components/layout.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/styles/components/table.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/lib/boxicons/boxicons.min.css" type="text/css" />
+   <link rel="stylesheet" href="/assets/lib/vanillajs-datepicker/datepicker.min.css" type="text/css" />

    <!-- JavaScript -->
-   <script defer src="/assets/js/index.js" type="module"></script>
+   <script defer src="/assets/lib/vanillajs-datepicker/datepicker-full.min.js" type="module"></script>
+   <script defer src="/assets/lib/vanillajs-datepicker/ko.js" type="module"></script>
+   <script defer src="/assets/lib/ckeditor5/ckeditor.js" type="module"></script>
+   <script defer src="/assets/lib/ckeditor5/ko.js" type="module"></script>
+   <script defer src="/assets/js/utils.js" type="module"></script>
+   <script defer src="/assets/js/common.js" type="module"></script>
</head>
```