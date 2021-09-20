using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practo_Web_Mini_Project
{
    class Program
    {
        public static void Main(string[] args)
        {
            //Creating instance for ExcelReadWriteCreate class
            ExcelReadWriteCreate Erw = new ExcelReadWriteCreate();

            //Calling function for writing data in excel
            Erw.WriteDataToExcel();

            Console.Read();
        }
        
    }
}
