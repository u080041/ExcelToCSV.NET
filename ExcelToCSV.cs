using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;


namespace TestSharepoint
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //Returns JSON "result":"success" if successful, or "Error: ..." otherwise
            //Example: [URL_PATH].aspx?etc_input=[INPUT_PATH]&etc_tab=[TAB_OF_EXCEL]&etc_ext=[XLS/XLSX]
            //Output will go into same folder as input


            // CHECK FOR VALID PARAMETERS SUPPLIED AND SAFE STATES
            //======================================================
            Regex pattern = new Regex(@"[^a-zA-Z0-9_:/@ \\]"); // limit valid inputs
            
            if (string.IsNullOrEmpty(Request["etc_input"]) || pattern.Match(Request["etc_input"]).Success==true)
            {
                Response.Write("{\"result\":\"Error: Invalid Input\"}");
                return;
            }
            else if (string.IsNullOrEmpty(Request["etc_tab"]) || pattern.Match(Request["etc_tab"]).Success == true)
            {
                Response.Write("{\"result\":\"Error: Invalid Tab\"}");
                return;
            }
            else if (string.IsNullOrEmpty(Request["etc_ext"]) || (Request["etc_ext"].ToLower() != "xlsx" && Request["etc_ext"]!="xls"))
            {
                Response.Write("{\"result\":\"Error: Invalid Extension\"}");
                return;
            }
            string reqTab = Request["etc_tab"],
                reqInput = Request["etc_input"]+"."+Request["etc_ext"],
                reqOutput = Request["etc_input"] + "(tab).csv";

            if (!File.Exists(reqInput))
            {
                Response.Write("{\"result\":\"Error: File '" + reqInput + "' does not exist\"}");
                return;
            }
            else if (File.Exists(reqOutput))
            {
                Response.Write("{\"result\":\"Error: File exists. Unable to overwrite file at  '" + reqOutput + "'.\"}");
                return;
            }



            // PROCESS EXCEL AND OUTPUT TO CSV
            //==================================
            // IMEX 1 to set type to Unicode Varchar, read in header row as content in csv
            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1;CharacterSet=UNICODE;\"", reqInput);

            string result = echoAsCSV(connectionString, reqTab);  
            if (result.StartsWith("Error:"))
            {
                Response.Write("{\"result\":\""+result+"\"}");
                return;
            }
            else
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(reqOutput);
                file.WriteLine(result);
                file.Close();
            }
            Response.Write("{\"result\":\"success\"}");
        }


        // CONVERSION METHOD [tab delimited]
        //====================================
        public static string echoAsCSV(string connectionString, String worksheetName)
        {
            string result = "";
            try
            {
                // Fill the dataset with information from the worksheet.
                var adapter1 = new OleDbDataAdapter("SELECT * FROM [" + worksheetName + "$]", connectionString);
                var ds = new DataSet();
                adapter1.Fill(ds, "results");
                DataTable data = ds.Tables["results"];

                // Process all rows and columns
                var lineResult = "";
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    lineResult = "";
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        if (j > 0)
                        {
                            lineResult += ("\t" + data.Rows[i].ItemArray[j]);
                        }
                        else
                        {
                            lineResult += (data.Rows[i].ItemArray[j]);
                        }
                    }
                    if (lineResult != "")
                    {                                                        // WILL NOT IGNORE BLANK LINES as long as got columns
                        if (result != "")
                        {
                            result += "\n" + lineResult;
                        }
                        else
                        {
                            result += lineResult;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("is not a valid name") > 0)
                {
                    result += "Error: Worksheet '" + worksheetName + "' may not be found in Workbook. Please Check. \n" + ex.Message.Replace("$", "");
                }
                else
                {
                    result += "Error: " + ex.Message;
                }
            }
            return result;
        }
    }
}
