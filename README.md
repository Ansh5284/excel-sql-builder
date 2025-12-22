# 📊 Excel SQL Builder (No-Code Data Connector)

> **Bridging the gap between Finance Teams and SQL Databases.**

![Status](https://img.shields.io/badge/Status-Prototype-orange) ![Stack](https://img.shields.io/badge/Tech-JavaScript%20%7C%20Office.js-blue)

## 🚀 The Problem
Financial Analysts often need data from company servers but lack the **SQL skills** to query it directly. They rely on IT departments for data dumps (CSV files), which leads to delays, version control issues, and manual copy-pasting errors.

## 💡 The Solution
**Excel SQL Builder** is a custom Office Web Add-in that brings a "Drag-and-Drop" interface directly into the Excel taskpane.
* **No Code Required:** Users select tables and columns visually.
* **Live Data:** Generates the SQL query in the background and pulls data directly into the active worksheet.
* **Secure:** Logic executes locally; credentials are handled via secure prompt (Phase 2).

![App Screenshot](./.github/screenshots/demo.png)
*(Screenshot of the Add-in Interface)*

## ⚙️ Key Features
* **Visual Query Builder:** Select columns, apply filters, and sort data using UI checkboxes.
* **Instant Export:** One-click fetch to populate the Excel grid.
* **Local Data Mode:** (Current V1) treats Excel Tables as databases for rapid analysis.

## 🛠️ Technical Architecture
* **Frontend:** HTML5, CSS3, Vanilla JavaScript.
* **Interaction:** Microsoft Office JavaScript API (Office.js) to manipulate Excel cells/ranges.
* **Logic:** Custom `SQLBuilder` class that constructs valid T-SQL syntax strings based on DOM selections.

## 📂 Project Structure
```text
src/
├── core/       # Logic for generating SQL strings
├── ui/         # HTML/CSS for the Taskpane
└── config/     # Connection string templates
