using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddComment
{
    public class AddComment : CodeActivity
    {
        [Category("输入")]
        [DisplayName("文件路径")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("输入")]
        [DisplayName("工作表名称")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; } = "Sheet1";

        [Category("输入")]
        [DisplayName("单元格")]
        [RequiredArgument]
        public InArgument<string> Cell { get; set; } = "A1";

        [Category("输入")]
        [DisplayName("批注")]
        [RequiredArgument]
        public InArgument<string> Comment { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            Excel.Sheets sheets = null;
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet workSheet = null;
            try
            {
                var filePath = FilePath.Get(context);
                var sheetName = SheetName.Get(context);
                var cell = Cell.Get(context);
                var comment = Comment.Get(context);
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(filePath);
                sheets = workbook.Sheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic val = sheets[i];
                    Excel.Worksheet worksheet = val as Excel.Worksheet;
                    if (worksheet != null && worksheet.Name.ToLowerInvariant().Equals(sheetName?.ToLowerInvariant()))
                    {
                        workSheet = worksheet;
                        workSheet.Activate();
                        break;
                    }
                    Marshal.ReleaseComObject(val);
                }
                Excel.Range oCell = workSheet.Range[cell, cell] as Excel.Range;
                if (oCell.Comment != null)
                {
                    oCell.Comment.Delete();
                }
                oCell.AddComment(comment);
                workbook.Save();
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workSheet);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
