using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practo_Web_Mini_Project
{
    public class TextFileWriteRead
    {
        public void DirectoryOperation()
        {
            string path = @"G:\Practo-Web-Mini_Project\creds.txt";

            if (File.Exists(path))
            {
                File.Delete(path);
                Console.WriteLine("File deleted");
            }

            using (StreamWriter sw = File.CreateText(path))
            {
                Console.WriteLine("New Text File Created!!!!!");
                sw.WriteLine("https://www.practo.com/doctors");
            }

            using (StreamReader rdr = File.OpenText(path))
            {
                Console.WriteLine("Reading from file");
                string message = string.Empty;
                Console.WriteLine(rdr.ReadToEnd());
            }

        }

        //

    }
}
