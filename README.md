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



## References
- Simplilearn Business Analytics with Excel Certification Training
- ChatGPT 4o



