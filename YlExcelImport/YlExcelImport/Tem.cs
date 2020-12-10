using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YlExcelImport.Models;

namespace YlExcelImport
{
   public class Tem
    {

        /// <summary>
        /// 获取excel配置对象
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public ExcelConfig GetExcelConfig(string filePath)
        {
            using (StreamReader srFile = File.OpenText(filePath))
            {
                using (JsonTextReader jsonReader = new JsonTextReader(srFile))
                {
                    JObject o = (JObject)JToken.ReadFrom(jsonReader);
                    var excelConfig = o.ToObject<ExcelConfig>();
                    return excelConfig;
                }
            }
        }

        /// <summary>
        /// ExcelConfig 文件校验
        /// </summary>
        /// <param name="excelConfig"></param>
        public void ExcelConfigCheck(ExcelConfig excelConfig)
        {
            if (excelConfig.ExcelType == 0)
            {
                throw new Exception("请配置json文件中的‘ExcelType’项，取值范围为：1，2，3");
            }
        }

        public void Te()
        {
            string path = "";
            var excelConfig = GetExcelConfig(path);
            switch (excelConfig.ExcelType)
            {
                case (int)ExcelTypeEnum.FormTable:
                    FormTable(excelConfig);
                    break;
                case (int)ExcelTypeEnum.NomalTable:
                    NomalTable(excelConfig);
                    break;
                case (int)ExcelTypeEnum.LevelTable:
                    LevelTable(excelConfig);
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 表单表格绘制
        /// </summary>
        /// <param name="excelConfig"></param>
        public HSSFWorkbook FormTable(ExcelConfig excelConfig)
        {

            var work = new HSSFWorkbook();
            var sheet= work.CreateSheet("121");
            var fields = excelConfig.FormHeader.FormFields;
            ExcelHelper excelHelper = new ExcelHelper(work, sheet, GetDt());
            excelHelper.CreateFormHeader(excelConfig.FormHeader);
            excelHelper.CreateTableHeader(excelConfig.TableHeader);
            return work;
        }

        public DataTable GetDt()
        {
            DataTable tblDatas = new DataTable("Datas");
            tblDatas.Columns.Add("Filed");
            tblDatas.Columns.Add("Value");

            tblDatas.Rows.Add(new object[] { "A", "地胜多负少方" });
            tblDatas.Rows.Add(new object[] { "B", "sd撒旦飞洒fs" });
            tblDatas.Rows.Add(new object[] { "C", "s士大夫a" });
            tblDatas.Rows.Add(new object[] { "D", "ad撒士fds"});
            return tblDatas;
        }

        public void NomalTable(ExcelConfig excelConfig)
        {

        }

        public void LevelTable(ExcelConfig excelConfig)
        {

        }






    }
}
