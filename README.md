# PS_Report_Spreadsheet

For this project, I looked at the column headers and data format of the csv report provided.  To maintain proper data security, I generated dummy data using mockaroo.com

The VBA code splits the data from the "Tags" column into two and trims off exteraneous text.  It also creates "Month" and "Quarter" columns based on the "Created At" column.

Next, all the data is copied to a named Table that the Pivot Tables are built from.  Slicers under the Pivot Tables allow the user to easily modify the parameters of the charts.

"PS Ticket Reports.xlsm" is the tool I created.  The other 3 csv files in the folder are different sets of dummy data.