# 📊 ScoreCard - Sales Analytics System

> **An internal enterprise mobile application built with .NET MAUI to automate sales reporting and data visualization.**

![C#](https://img.shields.io/badge/C%23-239120?style=for-the-badge&logo=c-sharp&logoColor=white)
![.NET MAUI](https://img.shields.io/badge/.NET_MAUI-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![Syncfusion](https://img.shields.io/badge/Syncfusion-UI_Controls-E94326?style=for-the-badge)
![Architecture](https://img.shields.io/badge/Architecture-MVVM-blue?style=for-the-badge)

## 📖 Overview

**ScoreCard** is a cross-platform application developed for **PowerHouse Data Centre** to streamline internal sales tracking. Before this app, sales reports required manual consolidation of multiple Excel files, which was time-consuming and prone to errors.

This solution automates the data ingestion process and presents key performance indicators (KPIs) through interactive dashboards, **reducing weekly reporting time by 70%**.

---

## 🏗️ Technical Architecture

The project follows the **MVVM (Model-View-ViewModel)** design pattern to ensure separation of concerns, testability, and maintainability.

### 📂 Project Structure (Based on Repository)

* **`ViewModels/`**: Contains the presentation logic and state management. Connects the UI to the data layer without direct dependencies.
* **`Views/`**: XAML-based UI pages (e.g., `DetailedSalesPage`) responsible for the visual layout.
* **`Models/`**: Defines data structures (e.g., `ProductSalesData`) and handles business objects.
* **`Services/`**: Handles data fetching and business logic processing (e.g., Excel parsing logic).
* **`Converters/`**: Custom value converters (e.g., `TotalVertivValueConverter`) for dynamic UI data binding.
* **`Controls/`**: Reusable UI components (e.g., `CustomDatePicker`) to maintain UI consistency across the app.

---

## ✨ Key Features

* **📈 Interactive Dashboards:** Utilizes **Syncfusion Charts** to visualize sales trends, regional performance, and product breakdowns.
* **🔄 Excel Automation:** Parses and consolidates raw Excel data automatically, eliminating manual data entry.
* **📱 Cross-Platform:** Built on **.NET MAUI** to run seamlessly on multiple platforms (iOS, Android, Windows) from a single codebase.
* **🔍 Detailed Analytics:** Provides drill-down capabilities for specific sales statuses and product categories.

---

## 🛠️ Tech Stack

| Category | Technologies |
|----------|--------------|
| **Framework** | .NET MAUI (Multi-platform App UI) |
| **Language** | C# |
| **Architecture** | MVVM (Model-View-ViewModel) |
| **UI Components** | Syncfusion, XAML |
| **Data Handling** | Excel Processing, LINQ |

---

## 🚀 Impact & Results

* **Efficiency:** Reduced weekly reporting time by **70%**.
* **Accuracy:** Eliminated human errors associated with manual data consolidation.
* **Decision Making:** Empowered management with real-time data visualization for faster strategic decisions.

---

## 📝 Note on Proprietary Code

*This project was developed during my internship at PowerHouse Data Centre. The full source code contains proprietary business logic and may not be fully public. This repository serves as a portfolio demonstration of the architecture and coding standards employed.*
