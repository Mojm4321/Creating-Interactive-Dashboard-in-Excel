# Creating-Interactive-Dashboard-in-Excel

##  Introduction
The first step in this process involved consolidating all relevant data into a single Excel worksheet in Excel to assess the overall performance of Databel Company. The dataset initially included daily and quarterly figures for total profit, customer growth and sales from January to September. It also contained metrics such as the average sales, profit and customer completion rate. Target sales were included to set performance benchmarks, enabling a comparison of actual performance against company standards. Additionally, the dataset incorporated regional data.

## Data Overview 
The data was presented in a plain table format (as seen in Diagram 1), which made it difficult to identify underlying patterns. To remedy this, headers were added (as seen in Diagram 2), to clearly distinguish each column, thereby reducing the risk of errors in large datasets. This allows for accurate referencing and analysis, adding a professional standard throughout the data analysis process.

![Diagram 1](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%201.png)
Diagram 1

![Diagram 2](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%202.png)
Diagram 2

## Data Breakdown
An additional column was introduced to reorganise the data month by month, (as seen in Diagram 3), providing a structured format alongside the ‘Date’ column. This is crucial because it maintains the professional feel to the data as most businesses frequently run on monthly cycles, and reduces the complexity associated with analysing daily activity. Additionally, colour coding was also employed (as seen in Diagram 3) to enhance the visual appeal and readability for those reviewing the analysis. 

![Diagram 3](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%203.png)
Diagram 3

Now that the table's visual presentation has been improved, naming it is essential (as shown in Diagram 4). Named tables faciliate easier data location, especially before developing the analysis and dashboard worksheets. Referring to the dataset name, rather than the cell ranges, reduces the chance of errors as the data is easier to manage. Diagram 5 illustrates how the Month column was created, including the formula used.

![Diagram 4](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%204.png)
Diagram 4

![Diagram 5](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%205.png)
Diagram 5

The next step involved creating a pivot table directly linked to the original dataset on a separate worksheet, as seen in Diagram 6. This worksheet is dedicated to summarising the data. This maintains data integrity and organisation within the workbook as datasets are easier to interpret when segregated into distinct sections. This is because the original dataset remains unchanged because analysis is done primarily with the pivot table going forward.  

 Furthermore, this setup automatically refreshes the pivot table with any changes made in the original dataset, guaranteeing that the conclusions drawn from the data are reliable and up to date without requiring manual adjustments every time new information is fed into the original dataset. 

![Diagram 6](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%206.png)
Diagram 6

As seen in Diagram 7, the pivot table selected the cumulative totals for profit, sales and customers across all months which provides a consolidated view of Databel’s overall performance. The pivot table automatically performs calculations like sums or averages, which eliminates the need for manual calculations, that can affects the accuracy of the data. At first glance, the figures indicate a strong finacial performance and customer reach.

![Diagram 7](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%207.png)
Diagram 7

When creating additional pivot tables to assess the performance of other metrics, you can simply copy and past the original table into an empty space on the workseet. Then, modify the metric in the field list to the one you need. This process is illustrated in both Diagram 7 & 8.

![Diagram 8](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%208.png)
Diagram 8

Diagram 9 shows all metrics have been copied, pasted and amended accordingly. The result of this process is displayed here, reflecting the updated metrics.

![Diagram 9](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%20%209.png)
Diagram 9

Diagram 10 demonstrates the calculation of sales and non-sales averages using the GETPIVOTDATA function in Excel. This function allows for precise extracton of specific data from the pivot tbale, which ensures accuracy in referencing data from the pivot tbale. By subtracting the original figures from 1, the non sales average is calculated, as the sales average is already provided.

![Diagram 10](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2010.png)
Diagram 10

## Dashboard Construction
Now that all these calculations have been completed for the key metrics, we can move on to building the dashboard. This key metrics selected are highlighted in Diagram 11.

![Diagram 11](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2011.png)
Diagram 11

The first step in constructing a dashboard involves creating an additional worksheet. Next, move to the view tab and remove both headings and gridlines to create an empty background for the dashboard, resembling Diagram 12. With this blank canvas, it is easier to position shapes that will show where the information is held, mirroring the layout depicted in Diagram 13.

![Diagram 12](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2012.png)
Diagram 12

![Diagram 13](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2013.png)
Diagram 13

It is important to orgainse metrics into distinct sections by assigning each shape a specific title. In this instance, the largest shape is designated as the 'Databel Dashboard', while smaller shapes are labeled with 'Sales', 'Profit', and 'Customers', as illustrated in Diagram 14 and 15.

Diagram 14 shows the method of entering text and numerical data into these shapes. Textboxes are inserted and the shape colour and outline is changed to white to blend with the green background. 

Once this setup is complete the same textbox is duplicated and formulas are applied to directly reference data from the Analysis worksheet, as seen in Diagram 14, '=Analysis!A4'. This approach avoids navigating between worksheets repeatedly, minmising the risk of errors. Diagram 15 shows the completed version on this process.

![Diagram 14](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2014.png)
Diagram 14

![Diagram 15](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2015.png)
Diagram 15

## Additonal Analysis
The next point of analysis involves creating pivot tables centered on completion rates, as shwon in Diagram 16. These completion rates serve as benchmark for evaluating sales productivity, profitability and customer management. By analysing these rates, Databel can identify its operations strengths and weaknesses, leading to more informed decision-making and targeted improvements to enhance its overall performance. 

To add the completion rate diagram to the dashboard, start by selecting the range containing the sales average and non sales average, as highlighted in Diagram 16. Then, go to the insert tab and choose the doughnut chart option. Once the chart is created, remove the title and the legend that displays 'sales average' and 'non sales average' at the bottom of the doughnut.Next, cut and paste the chart onto the dashboard worksheet. As shown in Digram 17, position the chart wihtin the designated shape. Use the "Format Data Series" option to adjust the size of the doughnut's inner part, which will visually represent the sales average figure.

![Diagram 16](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2016.png)
Diagram 16

![Diagram 17](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2017.png)
Diagram 17

Diagram 18 illustrates consistely high averages across all metrics. The sales average and non sales figures for profit and customers were added to their respective sections using the same steps for creating and formatting the doughnut charts.

![Diagram 18](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2018.png)
Diagram 18

First, duplicate the pivot table in the Analysis worksheet and adjust the field list options to display 'Target Sales' and 'Actual Sales', as shown in Diagram 19. Next, select the insert tab, click on the Recommended Charts section, choose the Column option and select the Stacked Column chart type. Customise the chart by hiding all field buttons, moving the legend to the bottom, as seen in Diagram 20. Once Diagram 20 has been created, duplicate it onto the Dashboard worksheet. Change the color theme to match the greeen background, and set the data labels to the 'Inside Base' option. These steps produce the chart seen in Diagram 21, which is well integrated into the dashboard with a consistent color theme and clear data labels.

![Diagram 19](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2019.png)
Diagram 19

![Diagram 20](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2020.png)
Diagram 20

![Diagram 21](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2021.png)
Diagram 21

Again, we begin by creating a pivot table in the Analysis worksheet to display the 'Sum of Customers' for each month, as shown in Diagram 22. Then, navigate to the insert tab and click on the recommended charts sections and select the line chart option, specifically the 'Stacked Line' chart. Customise the chart by removing the title and hiding all field buttons. After creating the chart in Diagram 22, duplicate it onto the Dashboard worksheet. Adjust the chart's colour scheme to blend with the background and change the chart line to ensure clear data presentation. This process results in Diagram 23, which is visually consistent with the dashboard. 

![Diagram 22](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2022.png)
Diagram 22

![Diagram 23](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2023.png)
Diagram 23

The final pivot table focused on the sum of sales by region. This was done by selecting 'Region' and 'Sales' from the field list, resulting in the table shown on the left in Diagram 24. Next, naviage to the insert tab, select, 'Recommended Charts' and choose the 'Clustered Bar' option to generate the chart displayed on the right in Diagram 24. To integrate this chart into the dashboard, removed the chart title and hide all field button. Copy and paste the chart into the dashboard worksheet and change the bar colours to match the green colour theme, as shown in Diagram 25. Images were added to the top right corner of the charts to enhance the visual appeal of the dashboard.

![Diagram 24](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2024.png)

Diagram 24

![Diagram 25](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2025.png)

Diagram 25

## Enhancing Interactivity
Slicers were added to make the dashboard more interactive, and to show which filters have been applied to the data. The slicers allow for multi select filerting, which allows for the data to analysed from various perspectives. The design looks fuller and fitted and synchorises the slicers with the rest of the dashboard. This can be seen in Diagram 26 and 27.

![Diagram 26](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2026.png)

Diagram 26

## Dashboard
![Diagram 27](https://github.com/Mojm4321/Creating-Interactive-Dashboard-in-Excel/blob/main/screenshots/Diagram%2027.png)

Diagram 27 

## Conclusion
The creative direction of datana analysis is critical for presenting information clearly and effectively. A well-designed dashboard minimises the risk of misinterpretation by utilising visual elements such as charts, graphs, pictures and colour-coding to highlight key insights and trends. This approach was prioritised to ensure Databel can thoroughly understand the data, as false interpretations could harm the sustainability of their business model through incorrect resource allocation or misguided financial planning. Data accuracy remains paramount; however, incorporating interactive features such as slicers allows the company to explore the data on specific areas, further enhancing their ability to make informed decisions.