# VBA_1
VBA program for downloading API records and displaying in Excel spreadsheets
			
1. Design workflow and algorithm			
	a. Connect and download the Hong Kong COVID-19 inflection cases data from Hong Kong Government website via API channel into MS Excel spreadsheets.		
	b. Apply VBA program code to extract, select and filter source data into Excel spreadsheet format.		
	c. Include error handling for exception cases.		
			
2. Program requirements			
	a. MS Excel 2012 or newer version installed.		
	b. Internet connected.		
	c. Macros enabled in MS Excel.		
	d. Enable content for external data connections. 		
	e. MS Excel file "HK_COVID_19_Cases_v1.xlsm" 		
			
3. Program File specification			
	a. Source Data Location worksheet		
		- API List : API source data paths and website locations	
	b. Download Records Worksheets		
		- Details of Cases : Hong Kong COVID-19 infection case records in detail.	
		- Confirmed Cases location: Hong Kong COVID-19 individual case records in districts.	
		- HK Population: Hong Kong 2019 population count in districts.	
	c. Result Display Worksheet		
		- Dashboard : three buttons	
			"Probable/Confirmed Cases Result" (select and filter the number of cases)
			"District Percentage Count Result" (select, filter and display no. of cases and percentage over population).
			"Source Data Download" (extract and download source data into worksheet "Details of Cases" and "Cases locations") *take approximately 3 minutes to complete the download due to the increase records.
	d. VBA codes in macros		
		- Refresh_Queries : function to refresh and extract two types of API data via HK gov website, linking up to button "Source Data Download" in worksheet "Dashboard”.	
			Cases_Detail_Download & Cases_Detail_Download_P: sub functions to download number of cases records
			Confirmed_Cases_Location_Download & Confirmed_Cases_Location_Download_P: sub functions to download district of probable/confirmed cases records
		- Count_Cases: function to select and display number of cases based on inputted start and end dates, linking up to button "Probable/Confirmed Cases Result" in worksheet "Dashboard”.	
		- Count_District_Pct: function to select and display number of cases in districts and percentage over population, linking up to button "District Percentage Count Result" in worksheet "Dashboard”.	
