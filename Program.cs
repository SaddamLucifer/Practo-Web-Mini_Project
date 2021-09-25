using Aspose.Cells;
using Aspose.Cells.Drawing;
//using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using System;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace Practo_Web_Mini_Project
{
    public class Program
    {
        public static List<String> ReadDataFromExcel(String path)
        {
            path = @"G:\Practo-Web-Mini_Project\Cities.xlsx";

            //Using NPOI pkg reading data from Excel

            XSSFWorkbook wb = new XSSFWorkbook(File.Open(path, FileMode.Open));
            XSSFSheet sh = (XSSFSheet)wb.GetSheetAt(0);
            XSSFRow row = (XSSFRow)sh.GetRow(0);
            XSSFCell cell = null;
            List<String> cell_values = new List<string>();
            int i, j;
            for (i = 1; i <= sh.LastRowNum; i++)
            {
                int cell_count = sh.GetRow(0).LastCellNum;
                for (j = 0; j < cell_count; j++)
                {
                    cell = (XSSFCell)sh.GetRow(i).GetCell(j);
                    String cell_value = cell.StringCellValue;
                    cell_values.Add(cell_value);
                }

            }
            return cell_values;
        }
        //**********Main Method*********************
        public static void Main(string[] args)
        {
            //
            string path = @"G:\Practo-Web-Mini_Project\Cities.xlsx";
            List<String> cell_values = ReadDataFromExcel(path);
            foreach (String cell_value in cell_values)
            {



                //Console.WriteLine(cell_value);
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
            //driver.FindElement(By.PartialLinkText("Book an app")).Click();

                      

            //Enter input into search box city
            driver.FindElement(By.CssSelector("input.c-omni-searchbox")).Clear();
            driver.FindElement(By.CssSelector("input.c-omni-searchbox")).SendKeys(Convert.ToString(cell_values[0]));
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//div[@data-qa-id='omni-suggestion-main' and text()='"+ cell_values[0]+ "']")).Click();
            Thread.Sleep(3000);
            /*ReadOnlyCollection<IWebElement> suggestions = driver.FindElements(By.XPath("//*[@id='c-omni-container']/div/div[2]/div[2]"));
            Thread.Sleep(3000);
            foreach (IWebElement suggestion in suggestions)
            {
                IWebElement hospital = driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[1]/div[2]/div[2]/div/span[1]"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(hospital).Click().Perform();
            }*/

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
            /*driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[1]/div/div[4]/span")).Click();
            ReadOnlyCollection<IWebElement> amenity = driver.FindElements(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[2]/div/div/div"));
            Thread.Sleep(2000);
            foreach (IWebElement amenities in amenity)
            {
                IWebElement pharmacy = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[2]/div/div/div/label[3]/div"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(pharmacy).Click().Perform();
            }

            Thread.Sleep(3000);*/

            //STEP 4-

            ReadOnlyCollection<IWebElement> ListOfHospitals = driver.FindElements(By.XPath("//*[@id='container']/div[3]/div/div[2]/div[1]/div/div[2]/div/div/div[1]/div[2]/div/div[1]/div/div/span[1]"));
            List<String> hospitals = new List<String>();
            for (int i = 2; i < ListOfHospitals.Count; i++)
            {
                IWebElement star_rating = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[2]/div[1]/div/div[2]/div[" + i + "]/div/div[1]/div[2]/div/div[1]/div/div/span[1]"));
                String[] stars = star_rating.Text.Split('.');
                int star_value = Int16.Parse(stars[0]);
                Console.WriteLine(star_value);
                if (star_value > 3)
                {
                    String hospital_name = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[2]/div[1]/div/div[2]/div[" + i + "]//h2")).Text;
                    Console.WriteLine(hospital_name);
                    
                    hospitals.Add(hospital_name);
                }

            }

            Console.WriteLine("Top 5 Hospitals for Search result are:");
            for (int i = 0; i < 5; i++)
            {     
                Console.WriteLine(cell_values[0]);
                Console.WriteLine(hospitals[i]);
            }


            //Console.WriteLine("Staus of Search Result for city Pune" + cell_values[0] + "Pass");


            Console.Read();
            Thread.Sleep(3000);
            driver.Quit();
        }

        
    }
}
