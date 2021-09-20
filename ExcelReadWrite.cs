using Aspose.Cells;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practo_Web_Mini_Project
{
    public class ExcelReadWrite
    {
        public void WriteDataToExcel()
        {
            //Creating array for storing city names in excel file
            string[] Cities = new string[7];
            
            Console.WriteLine("Enter Names of cities: ");

            //Accepting city names from user to write in excel file
            for (int i = 1; i < 7; i++)
            {
                Console.WriteLine("City-{0}: ", i);
                Cities[i] = Console.ReadLine();
            }
            
            //Path for Excel sheet to create and store
            string path = @"G:\Practo-Web-Mini_Project\Citie.xlsx";

            //Creating new Excel file and storing names of cities in it
            Workbook w = new Workbook();
            Worksheet sheet = w.Worksheets[0];
            w.Worksheets[0].Name = "City_Nmaes";  //Assigning name for sheet in excel file

            //Storing input city values to Excel file using for loop
            for (int j = 1; j < 7; j++)
            {
                //sheet.Cells[0][j].PutValue(Cities[j]);
                Cell cell = sheet.Cells["A" + j];
                cell.PutValue(Cities[j]);
            }
            w.Save(path, SaveFormat.Xlsx); //Saving Excel file in given path
            Console.WriteLine("Excel file created successfully...!!!");

        }
    }
}
