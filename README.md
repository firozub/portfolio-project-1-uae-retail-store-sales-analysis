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

Canvas dimensions were established at 1400 pixels width and 850 pixels height to match standard Power BI report proportions. Three distinct pages were planned for the dashboard. Page 1 Executive Summary was designated to focus on overall revenue trends, quantity moved, and most revenue contributing payment method across time. Page 2 Regional Analytics was designated to drill into geographic performance and profit trends and product category insights. Page 3 Data Quality and Cleanup was designated to showcase missing data to convey transparency and demonstrate the financial loss ie worth of missing data.

Five interactive slicers were designed and positioned on the left panel. Emirates and Region was created with multiple selection capability. Sales Channel was created as a dropdown filter to distinguish between sales channels. Payment Method was created as a dropdown filter to segment data by payment type. Year was created as a dropdown filter for temporal analysis. Product Category was created as a dropdown filter to drill into specific product lines.

The analytical approach for Page 1 Executive Summary was determined to be month-wise to show executive-level time trends and seasonal patterns throughout the year. Category-wise and regional analysis was reserved for subsequent analysis using quarterly and regional breakdowns.

Five analytical charts were planned for Page 1 Executive Summary sequenced to build business insights. A line chart displays total revenue by month to showcase monthly sales inflows and identify seasonal patterns and revenue momentum. A pie chart displays total revenue by payment method to reveal which payment method is most viable and preferred by customers and contributes most to revenue. A dual-line chart compares purchase price and sales price by month to visualize profit margin strategy and pricing decisions over time. A combination chart with columns for total revenue and line for quantity sold by month reveals whether revenue growth comes from increased volume or pricing decisions. A horizontal bar chart displays revenue by sales channel to compare online versus retail performance and identify the most effective distribution channel.

A Total Profit calculated column was created to enable profit analysis. The formula subtracted purchase price and discount cost from sale price multiplied by quantity sold to provide accurate profitability metrics. Info button descriptions were created for each page. Page 1 Executive Summary was assigned the title "Executive Performance Overview" with guidance on using filters and interpreting charts. Page 2 Regional Analytics was assigned the title "Regional Performance and Market Insights" with regional filtering guidance. Page 3 Data Quality and Cleanup was assigned the title "Data Quality and Processing Documentation" documenting the data cleaning process.

STEP 5: KPI DEFINITION & CALCULATIONS 

Four key performance indicators were defined for the Page 1 Executive Summary. These KPIs would be displayed as card visuals in the top row with associated growth percentages and visual indicators showing year-over-year performance trends.

Total Revenue was established as the primary revenue metric. The DAX formula created was `Total_Revenue_KPI = SUM([Total_Revenue])`. This measure sums all revenue generated from product sales. A companion growth measure was created to track year-over-year change with the formula:

Revenue_Growth =
VAR CurrentYearRevenue = SUM([Total_Revenue])
VAR PreviousYearRevenue = CALCULATE(SUM([Total_Revenue]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR _perc = DIVIDE(CurrentYearRevenue - PreviousYearRevenue, PreviousYearRevenue, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, UNICHAR(11165) & " " & FORMAT(_perc, "0.0%"),
_perc < 0, UNICHAR(11167) & " " & FORMAT(_perc*-1, "0.0%"),
FORMAT(_perc, "0.0%")
)

A conditional formatting measure was created to color-code the growth percentage with the formula:

CF_Revenue_Growth =
VAR CurrentYearRevenue = SUM([Total_Revenue])
VAR PreviousYearRevenue = CALCULATE(SUM([Total_Revenue]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR _perc = DIVIDE(CurrentYearRevenue - PreviousYearRevenue, PreviousYearRevenue, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, "Green",
_perc < 0, "Red",
"Grey"
)

Total Profit was established as the second primary KPI measuring profitability after all costs. The DAX formula created was `Total_Profit_KPI = SUM([Total Profit])`. A growth measure was created with the formula:

Profit_Growth =
VAR CurrentYearProfit = SUM([Total Profit])
VAR PreviousYearProfit = CALCULATE(SUM([Total Profit]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR _perc = DIVIDE(CurrentYearProfit - PreviousYearProfit, PreviousYearProfit, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, UNICHAR(11165) & " " & FORMAT(_perc, "0.0%"),
_perc < 0, UNICHAR(11167) & " " & FORMAT(_perc*-1, "0.0%"),
FORMAT(_perc, "0.0%")
)

A conditional formatting measure for profit growth was created with the formula:

CF_Profit_Growth =
VAR CurrentYearProfit = SUM([Total Profit])
VAR PreviousYearProfit = CALCULATE(SUM([Total Profit]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR _perc = DIVIDE(CurrentYearProfit - PreviousYearProfit, PreviousYearProfit, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, "Green",
_perc < 0, "Red",
"Grey"
)

Quantity Sold was established as the third KPI measuring sales volume and market movement. The DAX formula created was `Volume_of_Goods_Sold = SUM([quantity_sold])`. A growth measure was created to track volume changes year-over-year with the formula:

Volume_Growth =
VAR CurrentYearVolume = SUM([quantity_sold])
VAR PreviousYearVolume = CALCULATE(SUM([quantity_sold]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR _perc = DIVIDE(CurrentYearVolume - PreviousYearVolume, PreviousYearVolume, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, UNICHAR(11165) & " " & FORMAT(_perc, "0.0%"),
_perc < 0, UNICHAR(11167) & " " & FORMAT(_perc*-1, "0.0%"),
FORMAT(_perc, "0.0%")
)

A conditional formatting measure for volume growth was created with the formula:

CF_Volume_Growth =
VAR CurrentYearVolume = SUM([quantity_sold])
VAR PreviousYearVolume = CALCULATE(SUM([quantity_sold]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR _perc = DIVIDE(CurrentYearVolume - PreviousYearVolume, PreviousYearVolume, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, "Green",
_perc < 0, "Red",
"Grey"
)

Profit Margin Growth was established as the fourth KPI measuring profitability efficiency and margin trend. This measure compares the margin percentage of the current year against the previous year to show whether profitability per sale is improving or declining. The growth measure was created with the formula:

Profit_Margin_Growth =
VAR CurrentYearProfit = SUM([Total Profit])
VAR CurrentYearRevenue = SUM([Total_Revenue])
VAR CurrentYearMargin = DIVIDE(CurrentYearProfit, CurrentYearRevenue, 0)
VAR PreviousYearProfit = CALCULATE(SUM([Total Profit]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR PreviousYearRevenue = CALCULATE(SUM([Total_Revenue]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR PreviousYearMargin = DIVIDE(PreviousYearProfit, PreviousYearRevenue, 0)
VAR _perc = DIVIDE(CurrentYearMargin - PreviousYearMargin, PreviousYearMargin, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, UNICHAR(11165) & " " & FORMAT(_perc, "0.0%"),
_perc < 0, UNICHAR(11167) & " " & FORMAT(_perc*-1, "0.0%"),
FORMAT(_perc, "0.0%")
)

A conditional formatting measure for profit margin growth was created with the formula:

CF_Profit_Margin_Growth =
VAR CurrentYearProfit = SUM([Total Profit])
VAR CurrentYearRevenue = SUM([Total_Revenue])
VAR CurrentYearMargin = DIVIDE(CurrentYearProfit, CurrentYearRevenue, 0)
VAR PreviousYearProfit = CALCULATE(SUM([Total Profit]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR PreviousYearRevenue = CALCULATE(SUM([Total_Revenue]), YEAR('Electronics Retail Transaction Dataset'[Transaction_Date]) = YEAR(TODAY()) - 1)
VAR PreviousYearMargin = DIVIDE(PreviousYearProfit, PreviousYearRevenue, 0)
VAR _perc = DIVIDE(CurrentYearMargin - PreviousYearMargin, PreviousYearMargin, 0)
RETURN
SWITCH(
TRUE(),
_perc > 0, "Green",
_perc < 0, "Red",
"Grey"
)

All four KPI measures were tested with sample data across different slicer selections to ensure formula accuracy and proper functioning across filtered datasets. Each growth measure displays directional arrows and color coding to provide immediate visual feedback on performance trends.

Info button descriptions were created for Page 1 to guide users on dashboard purpose and usage. The color palette consisting of primary teal (#00838F) and secondary cyan (#00ACC1) and accent green (#4CAF50) was established for consistent visual theming across all KPI displays and chart elements.
