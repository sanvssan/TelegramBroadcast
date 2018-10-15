using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;
using System;

namespace telegramBroadcast
{
    class telegramapi
    {
        static void Main(string[] args)
        {
            excel.Application x1app = new excel.Application();
            excel.Workbook x1workbook = x1app.Workbooks.Open(@"C:\telegram\test.xlsx"); // Your excel file location for reading message
            excel._Worksheet x1worksheet = x1workbook.Sheets[2];
            excel.Range x1range = x1worksheet.UsedRange;

            string website;
            string telurl = "https://api.telegram.org/bot<Your api from bot father>/sendMessage?chat_id=";
            string broad;
            int maxsms;
            Console.WriteLine("Please enter number of message to broadcast to Telegram");
            maxsms = Convert.ToInt32(Console.ReadLine());
            for (int i = 1; i <= maxsms; i++)
            {
                website = x1range.Cells[i][1].value2;
                broad = telurl + website;
                Console.WriteLine(broad);
                IWebDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl(broad);
                Thread.Sleep(300);
                driver.Close();
                
            }
            x1workbook.Close();
            Environment.Exit(0);

        }
    }
}
