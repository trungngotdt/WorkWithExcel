using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkWithExecl
{
    class Program
    {
        static void Main(string[] args)
        {
            var address =Directory.GetParent( Directory.GetCurrentDirectory()).Parent.FullName.ToString()+ "\\Example.xlsx";
            ExcelReader excelReader = new ExcelReader(address);
            List<object> list= excelReader.GetData();
            
            list.ForEach((x) =>
            {
                Console.Write(x.ToString());
            });
            Console.ReadLine();
        }
    }
}
