# CoffeeSales-DataAnalysis

### Project Summary
This Excel project focuses on analysing and visualising coffee sales data for a fictional company. The original dataset was divided into three sheets: 'Orders,' 'Customers,' and 'Products.' However, crucial data was missing from the 'Orders' sheet. To address this, I employed advanced Excel techniques, such as XLOOKUP and INDEX MATCH, to populate the missing data from the other sheets.

After filling in the gaps, I performed several calculations to determine the order sales, standardised the data by transforming abbreviations into full names, and ensured proper formatting across all columns (e.g., converting numbers into currency and weights into kilograms). Finally, I created three pivot tables and visualisations to analyse the data from different perspectives and integrated them into an interactive dashboard.

### Techniques Utilised
<b>Data Lookup:</b>
<ul>
  <li><b>XLOOKUP: </b>Used to retrieve data from the 'Customers' and 'Products' sheets to populate the missing values in the 'Orders' sheet.</li>
  <li><b>INDEX MATCH: </b>Employed a dynamic formula to efficiently match and retrieve data across multiple columns, reducing the need for repetitive XLOOKUP formulas.</li>
</ul>

<b>Data Transformation & Formatting:</b>
<ul>
  <li>Cleaned and transformed abbreviations into full names for clarity.</li>
  <li>Applied appropriate formatting, such as converting numerical data into currency and weight measurements into kilograms.</li>
</ul>

<b>Data Analysis & Visualisation:</b>
<ul>
  <li><b>Pivot Tables</b>: Created three key pivot tables for:
    <ol>
      <li><b>Total Sales: </b>Analysed month-by-month sales for different coffee types, visualised through a line graph with interactive filters.</li>
      <li><b>Top 5 Customers: </b>Displayed a bar chart representing the top 5 customers by total sales.</li>
      <li><b>Sales by Country: </b>Visualised total sales by country using a bar chart.</li>
    </ol>
  </li>
  <li><b>Interactive Dashboard: </b>Combined the visualisations into a dynamic dashboard, linking filters (slicers) and timelines to ensure all elements reflect changes interactively.</li>  
</ul>

### Conclusion
This project has significantly enhanced my proficiency in Excel, particularly in the areas of data lookup, transformation, and visualisation. I learned the importance of optimising formulas for efficiency, such as using a dynamic INDEX MATCH formula with dual MATCH functions to streamline data retrieval across multiple columns, as opposed to unique XLOOKUP forumals for each column.

I continue to explore and learn new Excel techniques to improve my data analysis skills further. This project demonstrates my ability to tackle complex data challenges and create meaningful, interactive visualisations that provide valuable insights.
