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


STEP 5: KPI DEFINITION & CALCULATIONS 

Four key performance indicators were defined for the Page 1 Executive Summary. These KPIs would be displayed as card visuals in the top row with associated growth percentages and visual indicators showing year-over-year performance trends.

Total Revenue was established as the primary revenue metric. The DAX formula created was Total_Revenue_KPI = SUM([Total_Revenue]). This measure sums all revenue generated from product sales. A companion growth measure was created to track year-over-year change with the formula:

```dax
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
```

A conditional formatting measure was created to color-code the growth percentage with the formula:

```dax
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
```

Total Profit was established as the second primary KPI measuring profitability after all costs. The DAX formula created was Total_Profit_KPI = SUM([Total Profit]). A growth measure was created with the formula:

```dax
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
```

A conditional formatting measure for profit growth was created with the formula:

```dax
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
```

Quantity Sold was established as the third KPI measuring sales volume and market movement. The DAX formula created was Volume_of_Goods_Sold = SUM([quantity_sold]). A growth measure was created to track volume changes year-over-year with the formula:

```dax
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
```

A conditional formatting measure for volume growth was created with the formula:

```dax
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
```

Profit Margin Growth was established as the fourth KPI measuring profitability efficiency and margin trend. This measure compares the margin percentage of the current year against the previous year to show whether profitability per sale is improving or declining. The growth measure was created with the formula:

```dax
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
```

A conditional formatting measure for profit margin growth was created with the formula:

```dax
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
```

All four KPI measures were tested with sample data across different slicer selections to ensure formula accuracy and proper functioning across filtered datasets. Each growth measure displays directional arrows and color coding to provide immediate visual feedback on performance trends.

Two key performance indicator measures were created for Page 2 Regional Analytics to identify the most profitable region and its profitability characteristics. The Most_Profitable_Region and its Share of Total Profit measure was created using a complex DAX formula that identifies which region contributes the highest profit and displays that region's name with its share of total profit by calculating the maximum profit region and its percentage contribution to overall profit with the formula:

```dax
Most Profitable Region and its Share of Total Profit = 
VAR TotalProfit = SUM([Total Profit])
VAR MaxProfitByRegion = MAXX(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])))
VAR RegionWithMaxProfit = MINX(FILTER(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])) = MaxProfitByRegion), 'Electronics Retail Transaction Dataset'[Emirates / Region])
VAR MaxProfitPercent = DIVIDE(MaxProfitByRegion, TotalProfit, 0)
RETURN
    RegionWithMaxProfit & " " & FORMAT(MaxProfitPercent, "0.0%")
```

The Region_Profit_Margin_Pct measure was created to calculate the profit margin percentage of the most profitable region by dividing regional profit by regional revenue and multiplying by 100 and formatting with percentage display showing the profitability efficiency of the leading region with the formula:

```dax
Region Profit Margin % = 
VAR MaxProfitByRegion = MAXX(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])))
VAR RegionWithMaxProfit = MINX(FILTER(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])) = MaxProfitByRegion), 'Electronics Retail Transaction Dataset'[Emirates / Region])
VAR RegionProfit = CALCULATE(SUM([Total Profit]), 'Electronics Retail Transaction Dataset'[Emirates / Region] = RegionWithMaxProfit)
VAR RegionRevenue = CALCULATE(SUM([Total_Revenue]), 'Electronics Retail Transaction Dataset'[Emirates / Region] = RegionWithMaxProfit)
VAR Margin = DIVIDE(RegionProfit, RegionRevenue, 0) * 100
RETURN
    FORMAT(Margin, "0.00") & "%"
```

A conditional formatting measure CF_Most_Profit_Region was created to color-code the regional profitability metrics using a three-tier color system with the formula:

```dax
CF_Most_Profit_Region = 
VAR MaxProfitByRegion = MAXX(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])))
VAR RegionWithMaxProfit = MINX(FILTER(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])) = MaxProfitByRegion), 'Electronics Retail Transaction Dataset'[Emirates / Region])
VAR RegionProfit = CALCULATE(SUM([Total Profit]), 'Electronics Retail Transaction Dataset'[Emirates / Region] = RegionWithMaxProfit)
VAR RegionRevenue = CALCULATE(SUM([Total_Revenue]), 'Electronics Retail Transaction Dataset'[Emirates / Region] = RegionWithMaxProfit)
VAR Margin = DIVIDE(RegionProfit, RegionRevenue, 0) * 100
RETURN
    IF(Margin > 35, "Green", IF(Margin > 25, "Bright_Cyan", "Red"))
```

A conditional formatting measure CF_Region_Profit_Margin was created to color-code the profit margin percentage with the formula:

```dax
CF_Region_Profit_Margin = 
VAR MaxProfitByRegion = MAXX(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])))
VAR RegionWithMaxProfit = MINX(FILTER(ALLSELECTED('Electronics Retail Transaction Dataset'[Emirates / Region]), CALCULATE(SUM([Total Profit])) = MaxProfitByRegion), 'Electronics Retail Transaction Dataset'[Emirates / Region])
VAR RegionProfit = CALCULATE(SUM([Total Profit]), 'Electronics Retail Transaction Dataset'[Emirates / Region] = RegionWithMaxProfit)
VAR RegionRevenue = CALCULATE(SUM([Total_Revenue]), 'Electronics Retail Transaction Dataset'[Emirates / Region] = RegionWithMaxProfit)
VAR Margin = DIVIDE(RegionProfit, RegionRevenue, 0) * 100
RETURN
    IF(Margin > 35, "Green", IF(Margin > 25, "Bright_Cyan", "Red"))
```

Regional KPI measures display three-tier color coding where regions with profit margin exceeding 35 percent display green (#4CAF50) indicating excellent profitability, regions with profit margin between 25 and 35 percent display bright cyan (#00ACC1) indicating good profitability, and regions with profit margin below 25 percent display red (#FF5252) indicating profitability needing improvement. This color coding provides immediate visual feedback on regional profitability performance and enables executives to quickly identify regions requiring attention or presenting growth opportunities.

All KPI measures across both Page 1 and Page 2 were tested with sample data across different slicer selections to ensure formula accuracy and proper functioning across filtered datasets. Each measure displays appropriate visual indicators and color coding to provide immediate feedback on performance trends.




STEP 6: CHART VISUALIZATION & FINALIZATION 

All five analytical charts were created on Page 1 Executive Summary and configured with appropriate data fields and visual formatting. A line chart titled "Monthly Revenue Trends" was created displaying total revenue by month using transaction date on the x-axis grouped by month and total revenue on the y-axis to showcase monthly sales inflows and identify seasonal patterns. A pie chart titled "Revenue Contribution by Payment Method" was created displaying total revenue by payment method to reveal which payment method is most viable and preferred by customers and show contribution to overall revenue. A dual-line chart titled "Cost vs Selling Price Trends" was created comparing purchase price and sales price by month showing both lines on the same chart with month on the x-axis to visualize profit margin strategy and pricing decisions over time. A combination chart titled "Quantity Sold vs Revenue Trend" was created with columns for total revenue and a line for quantity sold by month on the same chart to reveal whether revenue growth comes from increased volume or from pricing decisions. A horizontal bar chart titled "Sales Channel Performance Comparison" was created displaying revenue by sales channel with sales channel on the y-axis and revenue on the x-axis to compare online versus retail performance and identify the most effective distribution channel.

Four analytical charts were created on Page 2 Regional Analytics for comprehensive regional insights. A custom shape map titled "UAE Regional Profit Intensity" was created by importing a TopoJSON file with emirates properly labeled and color-coded by profit intensity to show geographic performance distribution. A stacked column chart was created displaying total revenue by emirates and region with payment method shown as stacked segments to reveal payment method preferences and distribution across regions. A horizontal bar chart was created displaying cost of discount allowed by emirates and region with regions on the y-axis and discount cost on the x-axis to show discount allocation strategies and identify regions with highest discount burden. A multi-line chart was created displaying regional trends over time with quarterly revenue showing each region as a separate line across quarters to reveal quarterly performance trends and regional momentum patterns.

Two KPI callout cards were created on Page 2 to display regional profitability metrics. The Most_Profitable_Region callout card displays the region name with its percentage contribution to total profit and applies conditional formatting to color-code the metric based on profit levels. The Region_Profitability_Pct callout card displays the profit margin percentage of the most profitable region formatted with percentage display and applies conditional formatting to color-code based on margin thresholds.

Five interactive slicers were added to Page 1 Executive Summary positioned on the left panel. Emirates and Region slicer was created with radio button style for multiple selection capability enabling users to drill into specific geographic areas. Sales Channel slicer was created as a dropdown filter allowing users to isolate online or retail performance. Payment Method slicer was created as a dropdown filter to segment data by payment type. Year slicer was created as a dropdown filter for temporal analysis enabling year-to-year comparisons. Product Category slicer was created as a dropdown filter to drill into specific product lines and analyze category performance.

The same five interactive slicers were replicated on Page 2 Regional Analytics to maintain consistent filtering experience across pages. All slicers were configured to work across both pages enabling cross-page filtering when needed. Slicer styling was applied consistently with professional formatting and intuitive control layouts.

A cohesive color scheme was applied consistently across all visuals on both pages using a cold color palette of primary teal, secondary cyan, and turquoise tones to create visual harmony and professional appearance. All chart elements including axes, labels, and visual components were formatted using the established color palette to ensure consistency. Rounded borders and design polish were applied to all cards, KPI callouts, and chart containers to create a modern and professional aesthetic. Text formatting was standardized across all titles, labels, and descriptions maintaining readable font sizes and appropriate contrast ratios. Visual hierarchy was established through consistent use of colors, spacing, and sizing to guide user attention to key metrics and insights.

Info buttons and page descriptions were finalized on both pages. Page 1 Executive Summary was labeled with the title "Executive Performance Overview" and descriptive guidance text explaining the purpose and how to use filters. Page 2 Regional Analytics was labeled with the title "Regional Performance and Market Insights" and descriptive guidance text explaining regional analysis capabilities and filtering options. Company logo and branding elements were added to both page headers ensuring visual consistency and professional presentation throughout the dashboard.

Final testing was conducted across all charts visuals and slicers by applying various filter combinations to ensure proper data updates and cross-filtering functionality. KPI cards were verified to display correctly with growth percentages and color-coded indicators updating appropriately based on slicer selections. Map colors were tested to ensure proper color intensity representation across all emirates and regions. All tooltips were added to charts providing detailed information on hover. Dashboard responsiveness was verified to ensure proper display across different screen resolutions. The complete two-page dashboard was finalized with professional formatting and design polish ready for publication and presentation to stakeholders.
