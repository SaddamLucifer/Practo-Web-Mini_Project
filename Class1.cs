using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practo_Web_Mini_Project
{
    class Class1
    {
        /*my code
         driver.FindElement(By.CssSelector("input.c-omni-searchbox")).Clear();
            driver.FindElement(By.XPath("//input[@data-qa-id='omni-searchbox-locality']")).SendKeys(Convert.ToString(sheet.Cells["A1"].Value));
            Thread.Sleep(2000);


            IWebElement input_hospital = driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div/input"));
            input_hospital.SendKeys("hospital");
         */


        /*
          driver.FindElement(By.CssSelector("input.c-omni-searchbox")).Clear();
            driver.FindElement(By.XPath("//input[@data-qa-id='omni-searchbox-locality']")).SendKeys("Mumbai");
            Thread.Sleep(3000);

            driver.FindElement(By.XPath("//input[@data-qa-id='omni-searchbox-keyword']")).Clear();
            driver.FindElement(By.XPath("//*[@id='c-omni-container']/div/div[2]/div/input")).SendKeys("hospital");
            Thread.Sleep(3000);
     
            driver.FindElement(By.XPath("//div[@data-qa-id='omni-suggestion-main' and text()='Hospital']")).Click();
         */


        /*
           // STEP-3
           //Acredited
           IWebElement checkbox1 = driver.FindElement(By.Id("Accredited0"));
           checkbox1.Click();
           //24X7 Pharmacy
           IWebElement checkbox2 = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[2]/div/div/div/label[3]/div"));
           checkbox2.Click();
           */

        /*ReadOnlyCollection<IWebElement> ratings_hospitals = driver.FindElements(By.XPath("//div[@data-qa-id='hospital_card']//div[@data-qa-id='star_rating']//span[@class='common_star-rating_value']"));
        foreach (IWebElement rating in ratings_hospitals)
        {
            String[] star = rating.Text.Split(".");
            Console.WriteLine(star[0]);
            int star_value = Int16.Parse(star[0]);
        }
       */
        //last
        //IWebElement checkbox2 = driver.FindElement(By.XPath("//*[@id='container']/div[3]/div/div[1]/div/div/header/div[2]/div/div/div/label[3]/div"));
        //checkbox2.Click();


        /*ReadOnlyCollection<IWebElement> ratings_hospitals = driver.FindElements(By.XPath("//div[@data-qa-id='hospital_card']//div[@data-qa-id='star_rating']//span[@class='common_star-rating_value']"));
        foreach (IWebElement rating in ratings_hospitals)
        {
            String[] star = rating.Text.Split(".");
            Console.WriteLine(star[0]);
            int star_value = Int16.Parse(star[0]);
        }
       */

    }
}
//Creating instance for ExcelReadWriteCreate class
//ExcelReadWriteCreate Erw = new ExcelReadWriteCreate();
//Program pr = new Program();

/*pr.WriteDataToExcel();  //Calling function for writing data in excel

string path = @"G:\Practo-Web-Mini_Project\Cities.xlsx";
Worksheet sheet = null;

if (!File.Exists(path))
{
     sheet = pr.WriteDataToExcel();
}
*/