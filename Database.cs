using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace Test_Dushyant
{
    class Database
    {
        string consString = ConfigurationManager.ConnectionStrings["Test"].ConnectionString;

        /// <summary>
        /// Display data as per user query
        /// </summary>
        /// <param name="validFromDate"></param>
        /// <param name="commodityCode"></param>
        public void DisplayData(string validFromDate, string commodityCode)
        {
            using (SqlConnection con = new SqlConnection(consString))
            {
                SqlCommand command = new SqlCommand(ConfigurationManager.AppSettings["SearchQuery"], con);
                // Add the parameters
                command.Parameters.Add(new SqlParameter("vdate", validFromDate));
                command.Parameters.Add(new SqlParameter("cCode", commodityCode));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader[0] != null)
                    {
                        Console.WriteLine("Value fetched !!");
                        while (reader.Read())
                        {
                            Console.WriteLine(String.Format("{0} \t | {1} \t | {2} \t | {3} \t | {4} \t | {5} ",
                                reader[0], reader[1], reader[2], reader[3], reader[4], reader[5]));
                            Console.WriteLine("Data displayed!");
                            Console.ReadLine();
                        }
                    }
                    else
                    {
                        Console.WriteLine("No value return!! ");
                    }
                }
                Console.Clear();
            }

        }

        /// <summary>
        /// Load data from Data table to Database
        /// </summary>
        /// <param name="dtExcelData"></param>
        public void LoadData(DataTable dtExcelData)
        {
            using (SqlConnection con = new SqlConnection(consString))
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    //Set the database table name
                    sqlBulkCopy.DestinationTableName = "dbo.MyTest";

                    // Map the Excel columns with that of the database table
                    sqlBulkCopy.ColumnMappings.Add("Commodity Code", "CommodityCode");
                    sqlBulkCopy.ColumnMappings.Add("Diminishing Balance Contract", "DiminishingBalanceContract");
                    sqlBulkCopy.ColumnMappings.Add("Expiry Month Limit", "ExpiryMonthLimit");
                    sqlBulkCopy.ColumnMappings.Add("All Month Limit", "AllMonthLimit");
                    sqlBulkCopy.ColumnMappings.Add("Any One Month Limit", "AnyOneMonthLimit");
                    sqlBulkCopy.ColumnMappings.Add("Valid From", "ValidFrom");
                    con.Open();
                    sqlBulkCopy.WriteToServer(dtExcelData);
                    con.Close();
                }
            }
        }
    }
}
