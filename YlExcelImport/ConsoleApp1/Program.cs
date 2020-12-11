using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YlExcelImport;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Tem te = new Tem();
            var con = te.GetExcelConfig("ExcelConfig.json");
            var work = te.FormTable(con);
            using (FileStream url = File.OpenWrite(@"F:\练习\ExcelImport\Tes\Test.xls"))
            {
                //导出Excel文件
                work.Write(url);
            };
            Console.ReadLine();
        }
    }
}
