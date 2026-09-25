# UAE-Retail-Store-Sales-Analysis
UAE Retail Store Sales Analysis using Excel &amp; Power BI. Features data cleaning, KPI analysis, and interactive dashboards for profitability insights.

Project Overview

Analyzing 1000 rows of synthetic retail transaction data with 18% missing values using Power BI and Power Query Editor to demonstrate professional data cleaning, business insight generation, and interactive dashboard design.

Data Processing and Dashboard Creation Steps:

STEP 1: DATA IMPORT & INSPECTION 

Raw CSV data was uploaded into Power BI. All columns and their data types were reviewed to establish a baseline understanding of the dataset structure. The inspection revealed 18% missing values across various fields along with several format inconsistencies that required attention in subsequent blocks.

STEP 2: DATA CLEANING 

Blank rows were removed from transaction_id and customer_id columns as these are essential for transaction tracking. Duplicate rows in the purchase_id column were eliminated to ensure data accuracy. Error rows throughout the entire dataset were identified and removed to maintain data integrity and overall quality.

Missing values were handled strategically. Text fields including product category and region and sales channel and payment method were coded as "N/A" to preserve row count and enable proper segmentation analysis. Numerical fields were filled with 0 values. This decision was documented as follows: "Missing values in product category and region and payment method were coded as 'N/A' to preserve row count and enable segmentation analysis. ID columns with blanks were removed as they are essential for transaction tracking." The customer_id and purchase_id columns were removed entirely since they did not serve any analytical purpose. Transaction_id was retained for analysis purposes.

STEP 3: DATA FORMATTING & STANDARDIZATION 

The transaction date column presented an error when converting from text to date format. A custom formula was applied to resolve this issue. The formula Date.FromText([transaction_date], "en-US") instructed Power BI to correctly interpret the date values.

The discount column contained whole numbers rather than percentage values. A new custom column was created to convert the data properly to percentage format and the original column was removed.

All text fields were standardized by capitalizing the first letter of each word using the Text.Proper function. A misspelling was corrected where "Online" had been entered as "Oniline" in the sales channel field.

Two calculated columns were created to support financial analysis. Total Revenue was calculated as [sale_price] * [quantity_sold] to show actual revenue per transaction. Cost of Discount Allowed was calculated as ([sale_price] * [Discount_Percentage(%)] ) * [quantity_sold] to quantify profit lost to discounts. An average unit price column was not created because product types vary significantly and discount percentages are randomly distributed across transactions. This approach would not yield meaningful insights.
