<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&color=auto&height=80&section=header" width="100%" />
</p>

##### **English** | [한국어](README.ko.md)

# 📱 Automated Expense Logging System
> **iOS Shortcuts + Google Apps Script (GAS)**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
<a href="https://developer.mozilla.org/en-US/docs/Web/JavaScript"><img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=flat&logo=javascript&logoColor=black"></a>
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google-apps-script&logoColor=white)](https://developers.google.com/apps-script)
<a href="https://www.google.com/sheets/about/"><img src="https://img.shields.io/badge/Google_Sheets-34A853?style=flat&logo=google-sheets&logoColor=white"></a>
[![iOS Shortcuts](https://img.shields.io/badge/iOS%20Shortcuts-FF4500?style=flat&logo=shortcuts&logoColor=white)](https://support.apple.com/guide/shortcuts/welcome/ios)

An automated financial logging system that updates spending data to Google Sheets in seconds via iPhone Shortcuts and visualizes it through a real-time dashboard.

## 🚀 Key Features
- **One-Tap Logging**: Instantly input expenses via iOS Shortcuts without opening any apps.
- **Smart Parsing**: Automatically separates category and amount from natural language input (e.g., `Cafe 5000`).
- **Dynamic Dashboard**: Visualizes spending patterns with real-time Google Sheets charts.
- **Monthly Auto-Backup**: Features a serverless trigger that automatically archives monthly summaries and resets daily logs on the last day of each month.
- **Error Handling**: Supports correction via negative inputs (e.g., `-5000`) and handles data format exceptions.

## 🛠 Tech Stack
- **Client**: iOS Shortcuts
- **Backend**: Google Apps Script (JavaScript-based Serverless API)
- **Database/UI**: Google Sheets

## 🏗 System Architecture
1. **Input**: User enters "Item Amount" using an iPhone Shortcut.
2. **Transfer**: The Shortcut sends a `GET` request to the Google Web App URL.
3. **Processing**: GAS receives the data, parses it, and records it in the row corresponding to today's date.
4. **Maintenance**: On the last day of the month, a trigger moves the monthly total to the database and initializes the sheet.

## 📸 Screenshots
<img width="1710" height="681" alt="image" src="https://github.com/user-attachments/assets/50d2b0a1-e077-4c6e-bf44-7a932a3eb1df" />
<br>
<img width="1713" height="681" alt="image" src="https://github.com/user-attachments/assets/c053db7a-3ae0-49f1-81d6-fe513b103f60" />

## 💻 How to Setup
1. **Google Sheets**: Configure the sheet according to the provided template.
2. **Apps Script**: Copy the code from `src/code.gs` into your Google Sheets Apps Script editor and deploy as a Web App.
3. **iOS Shortcut**: Create a shortcut on your iPhone and connect it to your Web App URL.
4. **Action Button**: Assign the shortcut to the Action Button in iPhone settings (Available for iPhone 15 Pro and later).

<br>



---
© 2026 Seong-hun Bae.
