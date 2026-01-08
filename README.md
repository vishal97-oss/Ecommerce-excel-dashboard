# E-Commerce Excel Performance Dashboard & Automation

An interactive Excel-based business intelligence system for analyzing sales, inventory, and profitability in a home & kitchen e-commerce operation. This project combines dashboards, pivot analysis, optimization, and automation into one integrated decision-support tool.

---

## 📊 Project Overview

This Excel project models a real e-commerce business selling home and kitchen products across multiple regions. It provides:

- Real-time KPI monitoring  
- Inventory health tracking  
- Sales performance analysis  
- Profit optimization using Solver  
- Automated reporting and alerts using VBA  

All components are connected through structured tables, formulas, pivot tables, and automation controls.

---

## 📁 Workbook Structure

The Excel file contains the following major modules:

| Sheet | Purpose |
|------|--------|
| **Introduction** | Project overview, navigation buttons, and skill checklist |
| **Sales Data** | Transaction-level order data including revenue and profit |
| **Dashboard** | Visual KPI dashboard with charts and slicers |
| **Data Prep Sheet** | Inventory logic, KPI calculations, and automation formulas |
| **Pivot Analysis** | Revenue analysis by product, region, and category |
| **Optimization** | Goal Seek & Solver-based profit and capacity modeling |
| **Automation Controls** | Macro-driven buttons and inventory alerts |

---

## 📈 Dashboard & KPI System

The **Dashboard** provides a real-time overview of performance:

- Total Revenue  
- Total Profit  
- Units Sold  
- Revenue by Product  
- Revenue by Region  
- Category filters using slicers  

The dashboard updates automatically when data changes or when slicers are used.

---

## 📦 Inventory Intelligence

The **Data Prep Sheet** automatically evaluates inventory health using:

- **XLOOKUP / XMATCH**
- **IF / COUNTIF / COUNTIFS**
- **RANDBETWEEN**
- **AVERAGE**
- **IFERROR**

Each product is assigned a status:
- **OK** – stock above reorder level  
- **REORDER** – stock below threshold  

A KPI panel summarizes:
- Total products  
- Number needing reorder  
- Average stock levels  
- Average reorder thresholds  

---

## 📊 Pivot Analysis

The **Pivot Analysis** sheet allows management to view:

- Revenue by product  
- Revenue by region (East, West, North, South)  
- Revenue by category  

All pivots are slicer-enabled for fast filtering and comparison.

---

## 📉 Optimization (Goal Seek & Solver)

The **Optimization** sheet models business decisions such as:

- Units to sell  
- Price per unit  
- Cost per unit  
- Target profit  
- Production capacity  

Using **Excel Solver and Goal Seek**, the model calculates:
- Required units to meet profit targets  
- Whether production limits are binding  
- If revenue goals are satisfied  

This allows scenario planning and profit optimization.

---

## ⚙️ Automation & VBA Controls

The **Automation Controls** sheet includes:

- Macro-enabled buttons  
- Inventory alert triggers  
- Formatted report toggles  

These features automate:
- Inventory checks  
- Reporting logic  
- Visual alerts  

The system behaves like a small ERP-style control panel inside Excel.

---

## 📸 Screenshots

Visual screenshots of all major sheets (Dashboard, Pivot Tables, Optimization, Automation Controls, and Data Prep) are included in the `/screenshots` folder.

---

## 🛠 Tools & Techniques Used

- Excel Tables  
- Pivot Tables & Slicers  
- Advanced formulas (XLOOKUP, IF, COUNTIFS, IFERROR, AVERAGE)  
- Goal Seek  
- Solver  
- VBA Macros & Buttons  
- Dashboard charts  
- KPI modeling  

---

## 🎯 Purpose

This project demonstrates how Excel can be used as a **full business analytics platform** for:

- Financial tracking  
- Inventory management  
- Decision modeling  
- Executive dashboards  
- Automation  

It is designed to be both **exam-ready** and **industry-relevant**.

---

## 👤 Author

**Vishal Vaid**  
Excel for Business Analytics  
E-Commerce Performance & Automation System
