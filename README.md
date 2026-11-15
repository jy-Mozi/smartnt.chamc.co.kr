# BECARE 통합관리자
> 비플러스 헬스케어 화면 설계서의 P4 ~ P130에 해당

<br/>

## 📋 Table of Contents

- [🔖 About the Project](#-about-the-project)
    - [📎 URL](#-url)
- [💻 Getting Started](#-getting-started)
    - [⚙️ Installation](#-installation)
    - [🏃 Run Locally](#-run-locally)
    - [🚀 Deployment](#-deployment)
- [🗃️ Documentation](#-documentation)
    - [🌈 Components](#-components)
    - [📁 Directories](#-directories)
    - [📐 Layout](#-layout)
    - [🪄 library](#-libraries)

<br/>

## 🔖 About the Project

### 📎 URL
> 페이지별 URL은 [구글 시트](https://docs.google.com/spreadsheets/d/1jZ3yCCTQxcxdBSQB4jBGxy0hHYrVlTTJNX7toiGJl_8/edit?usp=sharing) 확인

| Deployment | URL                                                                    |
| ---------- | ---------------------------------------------------------------------- |
| `STAGE`    | [https://biocare-admin.netlify.app](https://biocare-admin.netlify.app) |

<br/>

## 💻 Getting Started
> 퍼블리싱 작업 시 코드 포맷팅 및 로컬 서버를 실행하기 위한 용도

### ⚙️ Installation
Install dependencies
```shell
$ npm install
```

### 🏃 Run Locally
Serve with hot reload at localhost:5500
```bash
$ npm run start
```

### 🚀 Deployment
❗️자동 배포❗
> `main` branch로 코드를 push 하면 Netlify를 통해 STAGE 서버에 자동으로 배포됩니다<br/>
> 안정적이고 테스트된 코드만 `main` branch로 push 하도록 주의하세요

<br/>

## 🗃️ Documentation

### 🌈 Components
> [UI Docs](https://biocare-admin.netlify.app/docs/components)에 정리된 코드 조각을 그대로 복사 붙여넣기해서 사용 가능

코드 복사 영역
```html
<details open>
  <summary>...</summary>
  <!-- 여기의 코드 조각 복사 -->
</details>
```

### 📁 Directories
```
biocare-admin
├── assets
├── docs
├── node_modules
├── pages
├── .gitignore
├── .prettierrc
├── package.json
├── package-lock.json
└── README.md
```
> 위 구조에서 `assets` 폴더와 `pages` 폴더의 코드로 개발 진행<br/>
> 그 외 폴더 및 파일은 개발에 직접적인 관련 ❌<br/>
> 디렉토리별 상세는 [여기](docs/directories.md)서 확인

### 📐 Layout
❗️`<head />` 변경 필요❗
> 퍼블리싱 작업을 효율적으로 진행하기 위해 CSS 및 JavaScript를 import 방식으로 적용했으나, 실제 개발 시에는 외부 파일 적용 방식으로의 변경 필요<br/>
> 외부 파일 적용 방법 및 레이아웃 관련 상세는 [여기](docs/layout.md)서 확인

### 🪄 Libraries
> 사용 라이브러리 상세는 [여기](docs/libraries.md)서 확인

**사용 라이브러리 목록**
- boxicons
- ckeditor5
- vanillajs-datepicker