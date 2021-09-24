using System;
using Aspose.Cells;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;

namespace MakeMyTrip_Project
{
   public class Program
    {
       static void Main(string[] args)
        {
           // WritingDataToExcel();

            ReadDataFromExcel();

            Console.WriteLine("hello");

            excel.Application excelSheet = new excel.Application();
            excel.Workbook w = excelSheet.Workbooks.Open(@"C:\Users\HP\source\repos\MakeMyTrip_Project\Cities.xlsx");
            excel.Worksheet sheet = excelSheet.Sheets[1];
            w.Worksheets[1].Name = "City Names";
            excel.Range range = sheet.UsedRange;



            //Read Url From Text File.

            string filepath = @"C:\Users\HP\source\repos\MakeMyTrip_Project\test.txt";
            List<string> lines = new List<string>();
           
            lines = File.ReadAllLines(filepath).ToList();
            foreach (string item in lines)
            {
                Console.WriteLine(item);
            }
            Console.Read();

            // Lunch ChromeDriver
            IWebDriver driver = new ChromeDriver(@"C:\Users\HP\Desktop\c# notes\selenium\chromedriver_win32");
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(80000);
            driver.Url = lines[0];
            Thread.Sleep(80000);
            Console.WriteLine("cell value...");

            //Source Input 

            driver.FindElement(By.Id("fromCity")).SendKeys(range.Cells[2].Value2);
            //Thread.Sleep(3000);
            driver.FindElement(By.XPath("//input[@placeholder='From']")).SendKeys(range.Cells[2].Value2);
            Thread.Sleep(3000);
            IWebElement fromcity = driver.FindElement(By.XPath("//ul[@role='listbox']/li[1]"));
            Console.WriteLine(fromcity.Text);
            Actions act = new Actions(driver);
            act.MoveToElement(fromcity).DoubleClick().Perform();
            //Thread.Sleep(3000);

            //Destination Input

            driver.FindElement(By.Id("toCity")).SendKeys(range.Cells[4].Value2);
            //Thread.Sleep(3000);
            driver.FindElement(By.XPath("//input[@placeholder='To']")).SendKeys(range.Cells[4].Value2);
            Thread.Sleep(3000);
            IWebElement tocity = driver.FindElement(By.XPath("//ul[@role='listbox']/li[1]"));
            //Thread.Sleep(3000);
            Console.WriteLine(tocity.Text);
            tocity.Click();
            driver.FindElement(By.TagName("body")).SendKeys(Keys.Tab);

            var dateTime = DateTime.Now;
            var date = dateTime.ToString("dd");
            int departureDate = Int16.Parse(date) + 2;
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//div[@class='dateInnerCell']/p[text()='" + departureDate + "']")).Click();
            Thread.Sleep(3000);
            Console.WriteLine("From And To CIties Selected Successfully");

            //Click To Search Btn
            driver.FindElement(By.LinkText("SEARCH")).Click();
            Thread.Sleep(30000);


            //Print Five Flights In Console And Text File

            ReadOnlyCollection<IWebElement> ListOfAirLinesName = driver.FindElements(By.XPath("//div[@class='listingCard']//span[contains(@class,'airlineName')]"));
            Thread.Sleep(30000);
           
           for (int i = 1; i <= 5; i++)
            {
                Console.Write("Result : "+ i + " " +  lines[i]);
                Console.WriteLine(ListOfAirLinesName[i].Text);
                lines.Add(ListOfAirLinesName[i].Text);
                Console.WriteLine("Added in Text file Successfully.");
            }
            File.WriteAllLines(filepath, lines);

            //Back To Home Page
           // driver.FindElement(By.XPath("//*[@id='root']/div/div[1]/div/div/div/div/nav/ul/li[1]/a")).Click();
            driver.Close();

        }

        //Read Data From Excel File
        private static void ReadDataFromExcel()
        {
            
            Console.WriteLine("Read value from excel");
            excel.Application excelSheet = new excel.Application();

            excel.Workbook w = excelSheet.Workbooks.Open(@"C:\Users\HP\source\repos\MakeMyTrip_Project\Cities.xlsx");
            excel.Worksheet sheet = excelSheet.Sheets[1];
            w.Worksheets[1].Name = "City Names";
            excel.Range range = sheet.UsedRange;

            

            string source1 = range.Cells[2].Value2;
            Console.WriteLine(source1);

            string destination1 = range.Cells[4].Value2;
            Console.WriteLine(destination1);

            string source2 = range.Cells[8].Value2;
            Console.WriteLine(source2);

            string destination2 = range.Cells[10].Value2;
            Console.WriteLine(destination2);

            string source3 = range.Cells[14].Value2;
            Console.WriteLine(source3);

            string destination3 = range.Cells[16].Value2;
            Console.WriteLine(destination3);



            Console.WriteLine("Read Data From Excel Successfully.");

           
            Console.ReadLine();
        }

        //Write Data To Excel Sheet
       /* private static void WritingDataToExcel()
        {
            Console.WriteLine("Enter 1st Source:");
            string source1 = Console.ReadLine();

            Console.WriteLine("Enter 1st Destination:");
            string destination1 = Console.ReadLine();

            Console.WriteLine("Enter 2nd Source:");
            string source2 = Console.ReadLine();

            Console.WriteLine("Enter 2nd Destination:");
            string destination2 = Console.ReadLine();

            Console.WriteLine("Enter 3rd Source:");
            string source3 = Console.ReadLine();

            Console.WriteLine("Enter 3rd Destination:");
            string destination3 = Console.ReadLine();


            string path = @"C:\Users\HP\source\repos\MakeMyTrip_Project\Cities.xlsx";
            Workbook w = new Workbook();
            Worksheet sheet = w.Worksheets[0];
            w.Worksheets[0].Name = "City Names";


           

            sheet.Cells["A1"].PutValue("Source: ");
            sheet.Cells["A2"].PutValue("Destination: ");

            sheet.Cells["A4"].PutValue("Source: ");
            sheet.Cells["A5"].PutValue("Destination: ");

            sheet.Cells["A7"].PutValue("Source: ");
            sheet.Cells["A8"].PutValue("Destination: ");


            sheet.Cells["B1"].PutValue(source1);
            sheet.Cells["B2"].PutValue(destination1);

            sheet.Cells["B4"].PutValue(source2);
            sheet.Cells["B5"].PutValue(destination2);

          
            sheet.Cells["B7"].PutValue(source3);
            sheet.Cells["B8"].PutValue(destination3);


            w.Save(path, SaveFormat.Xlsx);
            Console.WriteLine("Excel file created successfully...!!!");

            Console.ReadLine();
        }*/
     
    }
}
