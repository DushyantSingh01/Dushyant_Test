using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;

namespace Test_Dushyant
{
    public static class Helper
    {
        /// <summary>
        /// Validation of provided date for dd/mm/yyyy format
        /// </summary>
        /// <param name="vaildFromDate"></param>
        /// <returns></returns>
        public static bool DateValidation(string vaildFromDate)
        {
            Regex r = new Regex(@"\d{2}/\d{2}/\d{4}");

            if (r.IsMatch(vaildFromDate))
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }
        /// <summary>
        /// Load data of excel file in data table
        /// </summary>
        /// <param name="excelPath"></param>
       public static void LoadExcelMain(string excelPath)
        {

            string conString = string.Empty;
            string extension = Path.GetExtension(excelPath);

            // Checking extension of file
            switch (extension)
            {
                case ".xls":
                    conString = ConfigurationManager.ConnectionStrings["ODB1"].ConnectionString;
                    break;
                case ".xlsx":
                    conString = ConfigurationManager.ConnectionStrings["ODB"].ConnectionString;
                    break;

            }
            conString = string.Format(conString, excelPath);
            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {
                excel_con.Open();
                string testSheet = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["Test_Data"].ToString();
                DataTable dtExcelData = new DataTable();

                // To convert data to its actual type
                dtExcelData.Columns.AddRange(new DataColumn[6] { new DataColumn("Commodity Code", typeof(string)),
            new DataColumn("Diminishing Balance Contract", typeof(string)),
            new DataColumn("Expiry Month Limit",typeof(decimal)),
                new DataColumn("All Month Limit", typeof(int)),
                new DataColumn("Any One Month Limit", typeof(int)),
                new DataColumn("Valid From", typeof(string))});

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + testSheet + "]", excel_con))
                {
                    oda.Fill(dtExcelData);
                }
                excel_con.Close();
                Database ds = new Database();
                ds.LoadData(dtExcelData);
               
            }
        }
        /// <summary>
        /// Get user inputs
        /// </summary>
        /// <returns></returns>
       public static string GetUserInput()
        {
            Console.WriteLine("Enter Valid From Date in dd/mm/yyyy format: ");
            String validFromDate = Console.ReadLine();
            while (true)
            {
                // Checking in provided date is in dd/mm/yyyy format
                if (Helper.DateValidation(validFromDate) == true)
                    break;
                else
                {
                    Console.WriteLine("Please try again /n Enter Valid From Date in dd/mm/yyyy format: ");
                    validFromDate = Console.ReadLine();
                    continue;
                }
            }
            Console.WriteLine("Enter commodity code: ");
            string commodityCode = Console.ReadLine();
            Database ds = new Database();
            ds.DisplayData(validFromDate, commodityCode);

            Console.WriteLine("Do you want to view another record Y or N ");
            string userInput = Console.ReadLine();
            if (userInput == "Y" && userInput == "y") return "1";
            else return "0";

        }
    }
}
