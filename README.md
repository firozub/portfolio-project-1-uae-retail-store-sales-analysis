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

Three calculated columns were created to support financial analysis. Total Revenue was calculated as [sale_price] * [quantity_sold] to show actual revenue per transaction. Cost of Discount Allowed was calculated as ([sale_price] * [Discount_Percentage(%)] ) * [quantity_sold] to quantify profit lost to discounts. Total Profit was calculated as ([sale_price] - [purchase_price]) * [quantity_sold] - [Cost_of_Discount_Allowed] to show net profit after accounting for both purchase costs and discounts granted. An average unit price column was not created because product types vary significantly and discount percentages are randomly distributed across transactions. This approach would not yield meaningful insights.


STEP 4: DASHBOARD STRUCTURE PLANNING 

Canvas dimensions were established at 1400 pixels width and 850 pixels height to match standard Power BI report proportions. Two distinct pages were planned for the dashboard. Page 1 Executive Summary was designated to focus on overall revenue trends, quantity moved, and most revenue contributing payment method across time. Page 2 Regional Analytics was designated to drill into geographic performance and profit trends and product category insights.

Five interactive slicers were designed and positioned on the left panel. Emirates and Region was created with multiple selection capability. Sales Channel was created as a dropdown filter to distinguish between sales channels. Payment Method was created as a dropdown filter to segment data by payment type. Year was created as a dropdown filter for temporal analysis. Product Category was created as a dropdown filter to drill into specific product lines.

The analytical approach for Page 1 Executive Summary was determined to be month-wise to show executive-level time trends and seasonal patterns throughout the year. Category-wise and regional analysis was reserved for subsequent analysis using quarterly and regional breakdowns.

Five analytical charts were planned for Page 1 Executive Summary sequenced to build business insights. A line chart displays total revenue by month to showcase monthly sales inflows and identify seasonal patterns and revenue momentum. A pie chart displays total revenue by payment method to reveal which payment method is most viable and preferred by customers and contributes most to revenue. A dual-line chart compares purchase price and sales price by month to visualize profit margin strategy and pricing decisions over time. A combination chart with columns for total revenue and line for quantity sold by month reveals whether revenue growth comes from increased volume or pricing decisions. A horizontal bar chart displays revenue by sales channel to compare online versus retail performance and identify the most effective distribution channel.

A Total Profit calculated column was created to enable profit analysis. The formula subtracted purchase price and discount cost from sale price multiplied by quantity sold to provide accurate profitability metrics. Info button descriptions were created for each page. Page 1 Executive Summary was assigned the title "Executive Performance Overview" with guidance on using filters and interpreting charts. Page 2 Regional Analytics was assigned the title "Regional Performance and Market Insights" with regional filtering guidance.

Page 2 Regional Analytics was constructed by first establishing the foundational elements. The five interactive slicers from Page 1 were replicated to maintain consistent filtering capability across emirates and region, sales channel, payment method, year, and product category. A logo and branding elements were added to the page header for visual consistency. An info button was positioned on the page with the title "Regional Performance and Market Insights" and descriptive text stating "This page analyzes performance metrics by emirate and region including revenue contribution, discount allocation strategies, payment method preferences, and quarterly trends across current and previous years. Use filters to compare regional performance across sales channels, product categories, and payment methods to identify market opportunities and regional strengths."

The map visualization for Page 2 required a custom map implementation to accurately represent UAE emirates. The process began by visiting simplemaps.com and downloading a GeoJSON map file for the United Arab Emirates using country code AE. The downloaded GeoJSON file was processed through mapshaper.org to inspect and standardize region names. Each emirate name was verified by using the inspect features tool and hovering over each region to ensure names matched exactly with the dataset. Any special characters or spelling inconsistencies were corrected by clicking on regions, editing names, and removing extraneous characters. Once all region names were verified and corrected the map was exported as a TopoJSON file. The custom TopoJSON map file was then imported into Power BI by selecting the shape map visual and accessing map settings. The custom map option was selected and the previously exported TopoJSON file was uploaded replacing the default US map. The location field was populated with [Emirates / Region] establishing the geographic reference for each region on the custom map. The color saturation field was populated with [Total Profit] to visualize profit intensity across regions with color saturation ranging from light cyan for lowest profit to dark teal for highest profit regions. Critical attention was paid to ensuring exact name matching between custom map region names and Power BI dataset region names by cross-referencing the final TopoJSON file against the dataset to confirm identical spelling punctuation and formatting with no discrepancies.

Three additional analytical charts were designed and implemented for Page 2 to provide comprehensive regional insights beyond the map visualization. A stacked column chart displaying total revenue by emirates and region segmented by payment method was created using axis [Emirates / Region], legend [payment_method], and values SUM([Total_Revenue]) to reveal payment method preferences across regions. A horizontal bar chart displaying cost of discount allowed by emirates and region was created using axis [Emirates / Region] and value [Cost_of_Discount_Allowed] to show discount allocation strategies by geographic location. A multi-line chart displaying regional trends over time with quarterly revenue was created using axis [Transaction_Date] grouped by quarter, legend [region], and values SUM([Total_Revenue]) to reveal quarterly performance trends and regional momentum across the year with each region represented as a separate line.


