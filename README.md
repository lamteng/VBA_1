# VBA_1
VBA program for downloading API records and displaying in Excel spreadsheets
			
1. Design workflow and algorithm			
	a. Connect and download the Hong Kong COVID-19 inflection cases data from Hong Kong Government website via API channel into MS Excel spreadsheets.		
	b. Apply VBA program code to extract, select and filter source data into Excel spreadsheet format.		
	c. Include error handling for exception cases.		
			
2. Program specification and requirements			
	a. MS Excel 2012 or newer version installed.		
	b. Internet connected.		
	c. Macros enabled in MS Excel.		
	d. Enable content for external data connections. 
	e. MS Excel file "HK_COVID_19_Cases_v1.xlsm
			
3. Program File specification			
	a. Excel worksheets		
		- Dashboard : three buttons 	
			"Count Cases" (select and filter the number of cases)
			"Count District Percentage" (select, filter and display no. of cases and percentage over population).
			"Download source data" (extracting gov data) *take abit time to proceed the download
		- Details of Cases : Hong Kong COVID-19 infection case records in detail.	
		- HK Population: Hong Kong 2019 population count in districts.	
		- Confirmed Cases locations: Hong Kong COVID-19 individual cases records in districts.	
		- API List : API source data paths and website locations	
	b. VBA codes in macros		
		- Refresh_Queries : function to refresh and extract three API data via HK gov website, linking up to button "Refresh Source Data" in worksheet "Dashboard”.	
		- Count_Case: function to select and filter number of cases based on inputted start and end date, linking up to button "Count Cases" in worksheet "Dashboard”.	
		- Count_district: function to select and filter number of cases in districts and percentage over population, linking up to button "Count District Pct" in worksheet "Dashboard”.	
 			
