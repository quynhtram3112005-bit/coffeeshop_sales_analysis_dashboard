# Coffee Shop Sales Performance & Interactive VBA Dashboard

## 📌 Project Overview
This project delivers a comprehensive, interactive business intelligence solution designed to analyze coffee shop performance and customer purchasing behavior. Built entirely within **Microsoft Excel**, the project processes transactional data to track core KPIs, uncover operational trends, and deliver dynamic visualizations. 

Notably, this solution integrates **VBA Macros** to automate dashboard interactions, providing seamless chart-switching mechanisms and user-friendly navigation for non-technical stakeholders.

---

## 📂 Repository Structure
*   `data/`: Contains the database files.
*   `dashboard/`: Contains the final analytical product.
    *   `Coffe_sales.xlsm`: The core macro-enabled Excel workbook containing the raw data, processed Pivot Tables, VBA modules, and the final automated Dashboard.
*   `README.md`: Project documentation (this file).

---

## 🛠️ Key Features & Analytics Methodology

### 1. KPI Tracking & Financial Metrics
Designed dedicated executive Summary Cards to monitor business health at a glance:
*   **Total Revenue:** Tracking gross sales across different timelines and store branches.
*   **Transaction Volume:** Analyzing total foot traffic and order counts to measure store capacity.
*   **Average Transaction Value (ATV):** Measuring customer spending behavior per ticket to evaluate cross-selling and up-selling performance.

### 2. Multi-Dimensional Data Visualization
Developed dynamic Pivot Charts and data models to dissect operations through various lenses:
*   **Time-Series Trends:** Visualizing sales fluctuations by month, day of the week, and operational hours.
*   **Product Performance:** Segmenting revenue and quantities sold by product categories and specific SKUs.
*   **Payment Method Analysis:** Analyzing customer preferences (Cash, Credit Card, Digital Wallets) to optimize checkout workflows.

### 3. Advanced Interaction & VBA Automation
Elevated the standard Excel dashboard into a software-like application interface:
*   **Automated Chart Switching:** Programmed VBA macros to allow users to switch between different chart views (e.g., swapping from product category distribution to temporal trends) instantly via buttons.
*   **User-Friendly Navigation:** Implemented macro-driven buttons to clear filters, reset slicers, and navigate across workbook tabs effortlessly.
*   **Global Filtering:** Integrated **Slicers and Timelines** allowing stakeholders to drill down into specific months, locations, or product types dynamically.

<img width="696" height="460" alt="Screenshot 2026-06-15 at 16 01 45" src="https://github.com/user-attachments/assets/e8121d36-336c-41cd-b1cc-e252a9fa3e3f" />
<img width="713" height="468" alt="Screenshot 2026-06-15 at 16 02 04" src="https://github.com/user-attachments/assets/c3f34bfd-09d8-45dc-85a3-0d6883d98f01" />

---

## 💻 Tech Stack & Tools Used
*   **Software:** Microsoft Excel (Macro-Enabled Workbook - `.xlsm`)
*   **Core BI Tools:** Pivot Tables, Pivot Charts, Slicers, Interactive Timelines.
*   **Automation:** VBA (Visual Basic for Applications) for macro automation, UI controls, and chart switching.
*   **Analytical Functions:** Advanced logical and lookup formulas (`INDEX/MATCH`, `SUMIFS`, `AVERAGEIFS`, `IF`).


---

## 🚀 How to Explore the Project
* Open the file Coffe_sales.xlsm using Microsoft Excel.

* ⚠️ Important: Upon opening, click "Enable Macros" or "Enable Content" at the top yellow bar. This is required for the VBA chart-switching buttons and navigation automation to function.

* Interact with the Slicers and VBA navigation buttons on the main dashboard tab to filter data dynamically.
