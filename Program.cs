using Aspose.Cells;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Practo_Web_Mini_Project
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Creating instance for ExcelReadWriteCreate class
            ExcelReadWriteCreate Erw = new ExcelReadWriteCreate();

            //Erw.WriteDataToExcel();   //Calling function for writing data in excel

            //Program pr = new Program();
            string path = @"G:\Practo-Web-Mini_Project\Cities.xlsx";
            Worksheet sheet = null;

            if (!File.Exists(path))
            {
                 sheet = Erw.WriteDataToExcel();
            }
            //
            // pr.WriteDataToExcel();

            TextFileWriteRead Twr = new TextFileWriteRead();
            Twr.DirectoryOperation();

            //Launching Browser
            string creds = System.IO.File.ReadAllText(@"G:\Practo-Web-Mini_Project\creds.txt");
            Console.WriteLine(creds);
            string[] web_url = System.IO.File.ReadAllLines(@"G:\Practo-Web-Mini_Project\creds.txt");

            IWebDriver driver = new ChromeDriver();     //Download from Nudget Packages
            //IWebDriver driver = new FirefoxDriver();  //Download from Nudget Packages

            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20000);
            driver.Url = web_url[0];

            // Instantiate a Workbook object that represents Excel file.
            Workbook wb = new Workbook(path);

            // Access "Sheet" from the workbook.
            sheet = wb.Worksheets[0];

            //Enter input into search box city
            driver.FindElement(By.CssSelector("input.c-omni-searchbox")).Clear();
            
            driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[1]/div[1]/input")).SendKeys(Convert.ToString(sheet.Cells["A1"].Value));
            driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div[1]/input")).SendKeys("hospital");
            
            ReadOnlyCollection<IWebElement> suggestions = driver.FindElements(By.XPath("//*[@id='c-omni-container']/div/div[2]/div[2]"));
            Thread.Sleep(3000);
            foreach (IWebElement suggestion in suggestions)
            {
                IWebElement hospital = driver.FindElement(By.ClassName("c-omni-suggestion-item__content"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(hospital).Click().Perform();
                
            }


            //Console.Read();
            //driver.Quit();
        }


    }
}
