using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xml_To_Excel.Utility
{
    public class ExcelSelect
    {
        public string PathExelSelect { get; set; }
        public string SelectInExelFrom { get; set; }
        public string SelectInExelTo { get; set; }
    }

    public interface IExcelManager
    {
        Task<object[,]> ListExcelArrayMaker(ExcelSelect excelSelect);
    }

    public class ExcelManager : IExcelManager
    {
        public async Task<object[,]> ListExcelArrayMaker(ExcelSelect excelSelect)
        => await Task.Run(() =>
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;

            excel.Workbooks.Open(excelSelect.PathExelSelect);

            var exc = (object[,])excel
                .Range[excelSelect.SelectInExelFrom + ":"
                + excelSelect.SelectInExelTo].Value;

            excel.Quit();
            return exc;
        });
    }
}
