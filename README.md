C#.NET
======
.NET application for converting Excel to CSV.
* No need to define named range in Excel File. 
* Will convert just by specifying the excel filename, tab to convert, and file extension (e.g. xlsx).
* Will convert Excel to CSV even without Excel installed. 
  However, 2007 Office System Driver: Data Connectivity Components will be required instead. That is available free from http://www.microsoft.com/en-sg/download/details.aspx?id=23734
* Currently only runs on a Windows Machine. 



Usage
=====
[URL_PATH].aspx?etc_input=[INPUT_PATH]&etc_tab=[TAB_OF_EXCEL]&etc_ext=[XLS/XLSX]

Returns JSON {"result":"success"} if successful. Else {"result":"Error: ..."}



Requirements
============
* Runs on Windows Machine. Tested on Windows 7, .NET Framework 4.0
* If Microsoft Excel is not installed on Machine, 2007 Office System Driver: Data Connectivity Components can be used instead. 
  That is provided free by Microsoft @ http://www.microsoft.com/en-sg/download/details.aspx?id=23734
