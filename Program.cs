using System;
using System.Configuration;

namespace Test_Dushyant
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load excel data into DB
                Helper.LoadExcelMain(ConfigurationManager.AppSettings["FileLocation"]);

                Console.WriteLine("Data Uploaded into Databse, please enter to provide your input");
                Console.ReadLine();
                if (Helper.GetUserInput() == "0") return;
                else Helper.GetUserInput();
            }
            catch (Exception ex)
            {
                throw ex;
            }          
        }
        
       
    }
}
