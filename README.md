# PS_Report_Spreadsheet

I was asked to take report data on dev tickets created over the last year and create data visualization to help understand what issues are most prevelant.  The first Pivot Charts I made were fine, but it was clear there would be the need to re-make them monthly or every couple of weeks as time went on and new tickets continue to be created.  I decided to make this tool that can be used over and over, pasting in current report data at any time.

In order to develop this tool at home without taking any client data off work systems, I created dummy data using https://www.mockaroo.com/.  This site has built-in generators for many types of data.  For more customized values, I was able to write Ruby lambda functions and formulas to create the data I needed.  All data in this repository are procedurally generated rather than pulled from any production environment.

There are 3 visible Sheets and 3 hidden Sheets.  The first sheet is named "Start Here" and has instructions and 2 large buttons.  The first button deletes any report data left over from a previous use of the tool.  The second button runs several VBA subroutines to clean the data, copy it into a named Table, then refresh the Pivot Charts based on that Table.

The report data on the tickets had all ticket tag text crammed into a single column (each ticket has at least 2 tags).  The VBA code splits the data from the "Tags" column into two new columns and trims off exteraneous text (specifically "Primary Tag - " and "Secondary Tag - ").  The script also creates "Month" and "Quarter" columns based on the "Created At" column.

Slicers under the Pivot Tables allow the user to easily modify the parameters of the charts.

"PS Ticket Reports.xlsm" is the tool I created.  The 3 csv files in the "data" folder are different sets of generated dummy data.