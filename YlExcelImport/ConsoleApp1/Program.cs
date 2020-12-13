using ConsoleApp1.Test;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YlExcelImport;
using YlExcelImport.Models;

namespace ConsoleApp1
{

    class Tem
    {
        public string AA { get; set; }
        public string BB { get; set; }
        public string CC { get; set; }
        public string DD { get; set; }
        public string EE { get; set; }
        public string FF { get; set; }
        public string GG { get; set; }
    }
        class Program
    {
        static void Main(string[] args)
        {
            //ExcelYangL te = new ExcelYangL("../../ExcelConfig/CommonTableExcelConfig.json");
            //var work = te.ExcelCommonTableExprot(GetList());
            ExcelTest te = new ExcelTest("../../ExcelConfig/FormTableExcelConfig.json");
            var work = te.ExcelFormTableExprot(GetFormList(), GetList());
            using (FileStream url = File.OpenWrite(@"../../../../Tes/FormTableExcelConfig.xls"))
            {
                //导出Excel文件
                work.Write(url);
            };
            Console.ReadLine();
        }

        static List<Tem> GetList()
        {
            var list = new List<Tem>();
            for(var i = 0; i < 900; i++)
            {
                list.Add(new Tem { AA = "A_" + i.ToString(), BB = "B_" + i.ToString(), CC = "C_" + i.ToString(), DD = "D_" + i.ToString(), EE = "E_" + i.ToString(), FF = "F_" + i.ToString(),GG= "G_" + i.ToString() });
            }
            return list;
        }


        static List<FormOption> GetFormList()
        {
            var list = new List<FormOption>();
            for (var i = 1; i < 10; i++)
            {
                list.Add(new FormOption { FildName = "A_" + i.ToString(), FildValue="Value_"+i.ToString() });
            }
            return list;
        }
    }
}
