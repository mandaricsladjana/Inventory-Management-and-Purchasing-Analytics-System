# Excel-Inventory-Management-System

# 📦 Inventory Management & Purchasing Analytics System

### Excel | VBA | Inventory Control | Supply Chain Analytics | Purchasing Optimization

## 📌 Project Overview

The **Inventory Management & Purchasing Analytics System** is an Excel-based business solution designed to improve inventory visibility, support purchasing decisions, and streamline stock management through data analysis and automation.

The project combines inventory control principles with supply chain analytics to transform operational transaction data into actionable insights. It integrates stock monitoring, demand analysis, replenishment planning, purchasing recommendations, and reporting into a centralized Excel environment.

By combining Excel formulas, PivotTables, interactive reporting, and VBA automation, the system aims to reduce manual work, identify potential stock shortages, improve purchasing visibility, and support more informed, data-driven decisions.

The project simulates a real-world inventory management environment involving multiple warehouses, automotive parts, suppliers, stock movements, and purchasing requirements.

## 🎯 Business Problem

Inventory and purchasing teams must balance product availability with inventory costs while managing changing demand, supplier lead times, and stock movements across multiple locations.

When inventory information is maintained through disconnected spreadsheets or manual processes, several operational challenges can arise:

- **Stockout risks:** Insufficient inventory can disrupt operations and delay order fulfillment.
- **Excess inventory:** Overstocking ties up working capital and increases storage costs.
- **Manual purchasing decisions:** Replenishment decisions can become inconsistent without standardized calculations and clear recommendations.
- **Limited demand visibility:** Historical sales and stock movements may not be translated into useful demand indicators.
- **Supplier performance monitoring:** Purchasing teams need visibility into supplier activity and replenishment requirements.
- **Reporting inefficiencies:** Repetitive manual updates can consume time and increase the risk of errors.
- **Limited decision support:** Raw transaction records do not automatically provide actionable insights for inventory planning.

The purchasing and inventory teams need to answer questions such as:
- Which products are at risk of running out of stock?
- When should we reorder each product?
- How much should we purchase?
- How much money will the recommended orders require?
- Which products are overstocked or moving too slowly?
- How does demand influence purchasing requirements?
- How can we reduce the time spent manually updating inventory records and preparing reports?

### The Objective

Develop an integrated Excel solution that connects inventory records, transaction history, demand analysis, and purchasing requirements to support more efficient inventory control and data-driven supply chain decisions.

## 💡 Proposed Solution

The system uses structured Excel tables, calculation logic, and reporting tools to create a centralized inventory management workflow.

It is designed to help users:

1. Monitor current stock levels across warehouse locations.
2. Analyze historical inventory movements and product demand.
3. Identify products that require replenishment.
4. Calculate suggested order quantities and estimated purchasing costs.
5. Highlight inventory exceptions and potential stock risks.
6. Analyze inventory value and movement patterns.
7. Track purchasing activity and supplier-related information.
8. Automate data entry and improve the usability of the workbook through VBA.
9. Present operational KPIs and purchasing insights through reports and dashboards.
10. Explore historical trends to support demand planning and future purchasing decisions.

## ⚙️ Key Features

### 1. Inventory Monitoring & Stock Control

The inventory module provides a centralized overview of products across warehouse locations.

Planned and implemented capabilities include:

- Current stock monitoring.
- Product and supplier information.
- Inventory levels by warehouse.
- Reorder level monitoring.
- Stock position calculations.
- Inventory valuation.
- Low-stock and replenishment alerts.
- Identification of products requiring attention.

### 2. Demand Analysis & Replenishment Planning

Historical transactions are used to analyze stock movements and estimate replenishment requirements.

The replenishment logic incorporates indicators such as:

- Average daily demand.
- Supplier lead time.
- Safety stock.
- Reorder point.
- Current stock position.
- Suggested order quantity.
- Estimated order cost.

These calculations support purchasing decisions by helping identify when products may need replenishment and how much stock should be ordered.

The recommendations are intended to provide structured decision support rather than replace professional purchasing judgment.

### 3. Purchasing Analytics

The purchasing module connects inventory requirements with estimated ordering needs.

Its intended capabilities include:

- Identification of products requiring replenishment.
- Suggested purchasing quantities.
- Estimated purchasing expenditure.
- Supplier-level purchasing analysis.
- Comparison of inventory requirements across locations.
- Identification of excess or slow-moving inventory.
- Purchasing prioritization based on stock risk and replenishment requirements.

### 4. Transaction Management

A structured transaction table records inventory movements and provides the historical data required for inventory calculations and analysis.

The transaction model includes:

- Transaction ID.
- Transaction date.
- Product ID.
- Warehouse location.
- Quantity.
- Transaction type.

Transaction types include sales, receipts, and opening stock.

Sales are represented by negative quantities, while receipts and opening stock are represented by positive quantities. This convention allows stock movements to be analyzed consistently.

### 5. VBA-Based Data Entry & Automation

The project incorporates VBA to improve workbook usability and reduce repetitive manual entry.

The intended user interface includes:

- A data-entry form for recording transactions.
- Product search functionality.
- Transaction-type selection.
- Date and quantity inputs.
- Automated submission of transaction records to the relevant table.
- Validation of user inputs.
- A more streamlined workflow for recording inventory movements.

The objective is to make the workbook function more like a lightweight inventory management application than a collection of independent spreadsheets.

### 6. Reporting & KPI Analysis

The reporting component transforms inventory and transaction data into information that supports operational monitoring.

Potential KPIs include:

| KPI | Purpose |
|---|---|
| Total Inventory Value | Measures the financial value of available stock. |
| Total Units on Hand | Summarizes inventory quantities across locations. |
| Products Below Reorder Point | Identifies products that may require replenishment. |
| Estimated Purchasing Cost | Quantifies the estimated cost of suggested orders. |
| Average Daily Demand | Supports demand analysis and replenishment planning. |
| Inventory Turnover | Measures how frequently inventory is consumed or sold over a defined period. |
| Days of Inventory on Hand | Estimates how long current inventory may last at the observed demand rate. |
| Slow-Moving Inventory | Identifies products with limited movement over a defined period. |
| Supplier Purchasing Analysis | Provides visibility into purchasing requirements by supplier. |

The final KPI selection depends on the available data and the implementation of the relevant calculations.

### 7. Interactive Dashboards & Reporting

The workbook is designed to bring together inventory monitoring, purchasing analysis, and operational reporting.

Planned reporting features include:

- PivotTables for aggregating inventory and transaction data.
- Interactive slicers for filtering reports.
- Warehouse-level analysis.
- Supplier-level analysis.
- Inventory and purchasing summaries.
- Visual indicators for inventory exceptions.
- Navigation between workbook sections.

These features aim to make operational information easier to explore and interpret without manually reviewing individual transaction records.

### 8. Inventory Risk & Exception Management

The system aims to identify inventory situations that may require intervention.

Examples include:

- Products below their reorder point.
- Potential stockout situations.
- Products with excess inventory.
- Products with unusually low movement.
- Purchasing requirements that need prioritization.
- Inventory discrepancies that require investigation.

Exception-based monitoring helps users focus on the products and purchasing decisions that need attention rather than reviewing every inventory record individually.

### 9. Historical Trend Analysis & Predictive Insights

An additional development objective is to explore how historical inventory and demand data can support forward-looking purchasing decisions.

Potential extensions include:

- Historical demand trend analysis.
- Identification of changes in product movement.
- Demand forecasting.
- Comparison of historical demand with current inventory.
- Forecast-informed replenishment recommendations.

AI-assisted or predictive functionality will be evaluated as an extension of the core inventory management system.

Any forecasts will need to be validated against historical data before being used to guide real purchasing decisions.

## 🛠️ Tools & Technologies

| Technology | Application |
|---|---|
| Microsoft Excel 365 | Core application for inventory management and analysis. |
| Excel Tables | Structured storage of inventory and transaction data. |
| Excel Formulas | Stock calculations, demand indicators, and purchasing recommendations. |
| PivotTables | Aggregation and analysis of inventory and purchasing data. |
| PivotCharts & Slicers | Interactive reporting and filtering. |
| VBA | Data-entry forms, workflow automation, and user interaction. |
| Inventory Management Principles | Stock control, reorder points, safety stock, and replenishment planning. |
| Supply Chain Analytics | Inventory monitoring, purchasing analysis, and operational KPIs. |
| Predictive Analytics | Planned exploration of demand trends and forecasting. |

## 🗂️ Workbook Structure

The workbook is organized into dedicated sections for inventory management, transaction tracking, purchasing analysis, and reporting.

| Worksheet | Purpose |
|---|---|
| `Inventory` | Product master data, stock levels, inventory calculations, and replenishment recommendations. |
| `Transactions` | Historical inventory movements, including sales, receipts, and opening stock. |
| `Purchasing Analysis` | Purchasing requirements, estimated order costs, and replenishment analysis. |
| `Order` | Reorder information and purchasing-related records. |
| `Reports` | KPI summaries, PivotTables, and inventory reporting. |

The workbook also incorporates a VBA-based interface to support transaction entry and improve navigation.

## 📊 Analytical Approach

The project follows a structured analytical workflow:

**1. Data Organization**  
Structure inventory records, supplier information, warehouse locations, and transaction history into consistent tables.

**2. Data Validation**  
Check transaction types, quantities, product identifiers, dates, and relationships between inventory records and transaction history.

**3. Inventory Calculation**  
Calculate stock levels and derive inventory indicators from historical movements.

**4. Demand & Replenishment Analysis**  
Use historical demand, supplier lead times, and safety stock assumptions to estimate replenishment requirements.

**5. Purchasing Recommendations**  
Identify products requiring attention and estimate suggested order quantities and purchasing costs.

**6. Reporting & Visualization**  
Summarize operational data through KPIs, PivotTables, and interactive reporting.

**7. Automation & Decision Support**  
Use VBA and structured workflows to simplify transaction entry and make inventory insights more accessible.

## 📈 Business Value

The project is designed to demonstrate how an Excel-based analytical solution can support inventory and purchasing operations.

Potential business benefits include:

- Improved visibility into stock availability across warehouse locations.
- More consistent replenishment decisions.
- Earlier identification of potential stock shortages.
- Better visibility into excess and slow-moving inventory.
- More structured purchasing cost estimates.
- Reduced dependence on repetitive manual calculations.
- Improved accessibility of inventory and purchasing KPIs.
- A more standardized process for recording and analyzing stock movements.

The actual business impact will depend on the quality of the data, the accuracy of the inventory logic, and how the solution is used.

Quantitative improvements, such as time saved or reductions in stockout risk, should be measured through testing rather than assumed in advance.

## 🧪 Validation & Testing

The solution should be evaluated using realistic inventory scenarios, including:

- Products with sufficient stock.
- Products approaching their reorder point.
- Products below their reorder point.
- Products with no recent sales.
- Products with different supplier lead times.
- Products with varying demand levels.
- Receipts and sales recorded across different dates.
- Inventory movements across multiple warehouses.
- Changes in transaction data and their effect on reports.

Testing helps verify that stock calculations, replenishment recommendations, and reporting outputs remain consistent when the underlying data changes.

## 🚀 Future Development

Potential extensions include:

- Automated inventory alerts and exception notifications.
- Enhanced order tracking and purchasing status monitoring.
- Supplier delivery and order performance analysis.
- FIFO and LIFO inventory valuation models.
- Improved slow-moving and excess inventory analysis.
- Historical demand forecasting.
- AI-assisted purchasing insights.
- Additional dashboard visualizations.
- More comprehensive VBA automation.
- Quantitative measurement of process improvements.

These enhancements will be added and documented as they are developed and validated.

## 👩‍💻 Skills Demonstrated

This project brings together technical and business skills across several areas.

**Data Analytics**
- Data structuring and validation.
- KPI definition and calculation.
- Historical transaction analysis.
- Trend analysis.
- Reporting and visualization.
- Data-driven decision support.

**Microsoft Excel**
- Structured Excel Tables.
- Advanced formulas.
- Inventory calculations.
- PivotTables and slicers.
- Interactive reporting.
- Workbook organization.

**VBA & Automation**
- UserForm development.
- Data-entry workflow design.
- Input validation.
- Table interaction.
- Repetitive task automation.

**Inventory & Supply Chain Management**
- Inventory control.
- Stock movement tracking.
- Demand and replenishment analysis.
- Reorder point and safety stock calculations.
- Purchasing planning.
- Supplier analysis.
- Inventory risk monitoring.

**Business Analysis**
- Problem definition.
- Requirements translation.
- Process improvement.
- Exception management.
- Operational KPI development.
- Business-oriented analytical thinking.

## 💼 Project Context

This is a portfolio project developed to demonstrate how data analysis, inventory management principles, and Excel automation can be combined to address practical operational challenges.

The workbook uses a simulated inventory environment and is intended for educational and demonstration purposes. It does not represent the performance of a specific company.

## 👤 Author

**Sladjana Mandaric**

Quality Engineer | Data Analyst | Supply Chain & Business Analytics

This project reflects my interest in combining operational knowledge, data analysis, and process improvement to build practical solutions that support better business decisions.

---

*Project status: In development. Features and analytical capabilities will be updated as implementation and testing progress.*
