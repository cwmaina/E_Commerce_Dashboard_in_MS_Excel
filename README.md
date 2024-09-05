# E_Commerce_Dashboard_in_MS_Excel
A company wishes to add user control for product categories for customers to choose a category and view the trend month-by-month and product-by-product. 

## Table of Contents
- [Project Overview](#project-overview)
- [Data Sources](#data-sources)
- [Tools](#tools)
- [Data Cleaning](#data-cleaning)
- [Data Analysis](#data-analysis)
- [Results](#results)
- [References](#references)

## Project Overview
MS Excel was used to analyze sales based on product categories and create a sales dashboard that breaks down sales by product category.

## Data Sources
The primary dataset used for this analysis is the 'E_Commerce_Dashboard_Project.xlsx' file, from the Simplilearn Business Analytics with Excel Certification Training, containing detailed information of orders made.

## Tools
MS Excel - Data cleaning, Analysis and Visualisation [Download here](https://www.microsoft.com/en-au/microsoft-365/excel)

## Data Cleaning
In the initial data preparation phase, the following tasks were performed:

1. Data loading and inspection
2. Checking missing values

## Data Analysis
1. Create a Histogram for 'Shipping Days' i.e. Aging

- Select 'Data' tab
- Select 'Data Analysis'
- Select 'Histogram' and click 'Ok'
- In the Histogram dialogue box click the 'Labels' checkbox as there are labels in the data
- In the 'Input Reference' select 'SalesData!D1:D51291' and in the 'Bin Reference' select 'Working!K3:K7'
- In the 'Output' select a new worksheet for the binning table, click the histogram checkbox and then ok.

2. Prepare a table of sales and profit month wise in the working sheet

```MS Excel
Sales= SUMIFS('Sales Data'!$H:$H,'Sales Data'!$U:$U,Working!$B4,'Sales Data'!$F:$F,$R$3)
```
```MS Excel
Profit=SUMIFS('Sales Data'!$K:$K,'Sales Data'!$U:$U,Working!$B4,'Sales Data'!$F:$F,$R$3)
```

3. Create a user control combo box for the product category

- Select 'Developer'
- Select 'Insert'
- Select 'Form Controls' then 'Combo box'
- Draw the box on the Working Sheet
- Right click on the box and Select 'Format Control'
- 'Input Range' is the 'List of Product Categories' in Q2:Q5
- 'Cell Link' is $R$2
- 'Drop Down Lines' is 4

4. Create a Column Chart of the month wise table and region wise table

- Select 'Months' and 'Sales' columns
- Select 'Insert, Chart, Clustered Column'
- Select 'Region' and 'Sales' columns
- Select 'Insert, Chart, Clustered Column'

## Results
1. Histogram

![1_Histogram](https://github.com/user-attachments/assets/c166bd59-6075-4188-9fd3-ef207544c96b)

2. Table of sales and profit month wise in the working sheet filtered to 'Auto and Accessories'

![2_Table_of_Sales_and_Profit_by_Month](https://github.com/user-attachments/assets/8d156110-28b8-4030-b93e-f181113558de)

3.  User control combo box for the product category filtered to 'Auto and Accessories'

![3_User_Control_Combo_Box](https://github.com/user-attachments/assets/036f1660-2465-4087-930b-828b56344a6b)

4. Column Chart of the month-wise table and region-wise table filtered to 'Auto and Accessories'

![4_Column_Chart_Month_Wise](https://github.com/user-attachments/assets/6508de55-605f-4d48-9509-7b5210bc7660)

![4_Column_Chart_Region_Wise](https://github.com/user-attachments/assets/e35a25fd-a3f9-4b05-b46d-889142784e93)


## References
- Simplilearn Business Analytics with Excel Certification Training
- ChatGPT 4o



