using Aspose.Cells;
using Aspose.Cells.Drawing;
//using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
//using System.Text;
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

            string path = @"G:\Practo-Web-Mini_Project\Cities.xlsx";
            Worksheet sheet = null;

            if (!File.Exists(path))
            {
                 sheet = Erw.WriteDataToExcel();
            }
            

            TextFileWriteRead Twr = new TextFileWriteRead();
            Twr.DirectoryOperation();

            //Launching Browser
            string creds = System.IO.File.ReadAllText(@"G:\Practo-Web-Mini_Project\creds.txt");
            Console.WriteLine(creds);
            string[] web_url = System.IO.File.ReadAllLines(@"G:\Practo-Web-Mini_Project\creds.txt");

            IWebDriver driver = new ChromeDriver();     //Download from Nudget Packages
            //IWebDriver driver = new FirefoxDriver();  //Download from Nudget Packages

            driver.Manage().Window.Maximize();
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10000);
            driver.Url = web_url[0];
            Thread.Sleep(3000);
            driver.FindElement(By.PartialLinkText("Book an app")).Click();

            // Instantiate a Workbook object that represents Excel file.
            Workbook wb = new Workbook(path);

            // Access "Sheet" from the workbook.
            sheet = wb.Worksheets[0];

         

            //Enter input into search box city
            driver.FindElement(By.CssSelector("input.c-omni-searchbox")).Clear();
            driver.FindElement(By.CssSelector("input.c-omni-searchbox")).SendKeys(Convert.ToString(sheet.Cells["A1"].Value));

            driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div/input")).Clear();
            driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div/input")).SendKeys("hospital");
            Thread.Sleep(3000);

            //driver.FindElement(By.XPath("//div[@data-qa-id='omni-suggestion-main' and text()='Hospital']")).Click();
            
            ReadOnlyCollection<IWebElement> suggestions = driver.FindElements(By.XPath("//*[@id='c-omni-container']/div/div[2]/div[2]"));
            Thread.Sleep(3000);
            foreach (IWebElement suggestion in suggestions)
            {
                IWebElement hospital = driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div[2]/div[1]/div[1]/span[1]/div"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(hospital).Click().Perform();
            }



            // STEP-3
            //Acredited
            driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[1]/div")).Click();
            ReadOnlyCollection <IWebElement> acccredited = driver.FindElements(By.XPath("/html/body/div[1]/div/div[3]/div/div[1]/div/div/header/div[1]/div/div[2]/label/div"));
            Thread.Sleep(2000);
            foreach (IWebElement option in acccredited)
            {
                IWebElement acc = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[1]/div/div[2]/label/div"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(acc).Click().Perform();
            }

            //IWebElement checkbox1 = driver.FindElement(By.Id("Accredited0"));
            //checkbox1.Click();
            Thread.Sleep(2000);

            //24X7 Pharmacy
            driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[1]/div/div[4]/span")).Click();
            ReadOnlyCollection<IWebElement> amenity = driver.FindElements(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[2]/div/div/div"));
            Thread.Sleep(2000);
            foreach (IWebElement amenities in amenity)
            {
                IWebElement pharmacy = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[2]/div/div/div/label[3]/div"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(pharmacy).Click().Perform();
            }

            Thread.Sleep(3000);

            //STEP 4-
            ReadOnlyCollection<IWebElement> ratings_hospitals = driver.FindElements(By.XPath("//div[@data-qa-id='hospital_card']//div[@data-qa-id='star_rating']//span[@class='common_star-rating_value']"));
            foreach (IWebElement rating in ratings_hospitals)
            {
                string[] star = rating.Text.Split('.');
                Console.WriteLine(star[0]);
                int star_value = Int16.Parse(star[0]);
                if (star_value > 3)
                {
                    /*
                    ReadOnlyCollection<IWebElement> top_hospitals_list = driver.FindElements(By.XPath("//div[@class='']//h2[data-qa-id*='hospital_name']"));
                    Thread.Sleep(3000);
                    foreach (IWebElement list in top_hospitals_list)
                    {
                        string List= list.Text;
                        Console.WriteLine(List);
                    }
                    */
                }
            }



            Console.Read();
            Thread.Sleep(3000);
            driver.Quit();
        }


    }
}
