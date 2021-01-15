using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using QiHe.CodeLib;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ganss.Excel;

namespace Introduction_ConsoleApp_Excel
{
    class Excel1
    {
        private static void Main(string[] args)
        {

            //strings for first and last anme
            string FirstName;
            string LastName;


            //Taking the information and writing to a file
            Console.WriteLine("Enter your first name");
            FirstName = Console.ReadLine();
            Console.WriteLine("Enter your last name");
            LastName = Console.ReadLine();




            //Display information to the console
            Console.WriteLine("Hello" + ' ' + FirstName + ' ' + LastName);





            //Writing First and Last name into the file

            StreamWriter sw = new StreamWriter("c:/users/leagu/source/repos/Introduction-ConsoleApp-Excel/ReadandWrite.txt");
            sw.WriteLine(FirstName + ' ' + LastName);
            sw.Close();
            //close the file



            // Creating a new List for name
            var students = new List<NAME>
            { //getting the first and last name input from above
                new NAME{FirstName=FirstName},
                new NAME{LastName=LastName}
            };
            //using ExcelMapper and creating a filename for XLSX
            ExcelMapper mapper = new ExcelMapper();
            var newFile = ("c:/users/leagu/source/repos/Introduction-ConsoleApp-Excel/ReadandWrite.xlsx");
            mapper.Save(newFile, students, "SheerName", true);
            Console.ReadKey();
        
        }
        //Creating class for NAME and getting the first and last name
        public class NAME
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }
    }
}
