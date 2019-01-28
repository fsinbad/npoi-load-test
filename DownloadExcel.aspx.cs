using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DownloadXlsx
{
    public partial class DownloadExcel : System.Web.UI.Page
    {
        

        protected void Page_Load(object sender, EventArgs e)
        {
            Button1_Click(sender, e);
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string filename = "test.xlsx";

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));

            XSSFWorkbook workbook = new XSSFWorkbook();
            Stopwatch sw = new Stopwatch();
            sw.Start();

            for (int i = 1; i <= 100; i++)
            {

           

             
                ISheet sheet1 = workbook.CreateSheet("Sheet"+i);

            //sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            //int x = 1;
            //for (int i = 1; i <= 15; i++)
            //{
            //    IRow row = sheet1.CreateRow(i);
            //    for (int j = 0; j < 15; j++)
            //    {
            //        row.CreateCell(j).SetCellValue(x++);
            //    }
            //}


                for (int rownum = 0; rownum < 100; rownum++)
                {
                    IRow row = sheet1.CreateRow(rownum);
                    for (int celnum = 0; celnum < 20; celnum++)
                    {
                        ICell Cell = row.CreateCell(celnum);
                        Cell.SetCellValue("Cell: Row-" + rownum + ";CellNo:" + celnum);
                    }
                }
            }

            sw.Stop();
            System.Console.WriteLine("cost " + sw.Elapsed.TotalSeconds + " s");
            System.Diagnostics.Debug.WriteLine("cost " + sw.Elapsed.TotalSeconds + " s");
            string tempFile = System.IO.Path.GetTempFileName();
            using (var f = File.Create(tempFile))
            {
                workbook.Write(f);
            }
            Response.WriteFile(tempFile);
            //http://social.msdn.microsoft.com/Forums/en-US/3a7bdd79-f926-4a5e-bcb0-ef81b6c09dcf/responseoutputstreamwrite-writes-all-but-insetrs-a-char-every-64k?forum=ncl
            //workbook.Write(Response.OutputStream); cannot be used 
            //root cause: Response.OutputStream will insert unnecessary byte into the response bytes.
            Response.Flush();
            Response.End();
        }
    }
}