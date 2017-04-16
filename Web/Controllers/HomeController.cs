using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Web.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            FormatReportWord();
            //ReportWord();
            return View();
        }
        private void FormatReportWord()
        {
            try
            {
                string templateFile = Server.MapPath("~/Temp/FormartWordTemplate.doc");
                string saveDocFile = Server.MapPath("~/Word/FormartWord" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc");
                Aspose.Words.Document doc = new Aspose.Words.Document(templateFile);
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                builder.MoveToBookmark("模块名称");
                //builder.InsertFootnote(0,"这是一个模块");
                builder.Write("这是一个模块");
                builder.MoveToBookmark("操作动作");
                //builder.InsertFootnote(0,"这是一个模块");
                builder.Write("添加一个门店");
                doc.Save(saveDocFile);
                System.Diagnostics.Process.Start(saveDocFile);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        private void ReportWord()
        {
            try
            {
                string templateFile = Server.MapPath("~/Temp/WordTemplate.doc");
                string saveDocFile = Server.MapPath("~/Word/" + DateTime.Now.ToString("yyyyMMddHHmmss")+".doc");
                Aspose.Words.Document doc = new Aspose.Words.Document(templateFile);
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                DataTable nameList = new DataTable();
                 
                nameList.Columns.Add("编号",typeof(string));
                nameList.Columns.Add("姓名", typeof(string));
                nameList.Columns.Add("时间", typeof(string));
                DataRow row = null;
                for (int i = 0; i < 50; i++)
                {
                    row = nameList.NewRow();
                    row["编号"] = i.ToString().PadLeft(4, '0');
                    row["姓名"] = "伍华聪 " + i.ToString();
                    row["时间"] = DateTime.Now.ToString();
                    nameList.Rows.Add(row);
                }

                List<double> widthList = new List<double>();
                for (int i = 0; i < nameList.Columns.Count; i++)
                {
                    builder.MoveToCell(0, 0, i, 0); //移动单元格
                    double width = builder.CellFormat.Width;//获取单元格宽度
                    widthList.Add(width);
                }

                builder.MoveToBookmark("table");        //开始添加值
                for (var i = 0; i < nameList.Rows.Count; i++)
                {
                    for (var j = 0; j < nameList.Columns.Count; j++)
                    {
                        builder.InsertCell();// 添加一个单元格                    
                        builder.CellFormat.Borders.LineStyle = LineStyle.Single;
                        builder.CellFormat.Borders.Color = System.Drawing.Color.Black;
                        builder.CellFormat.Width = widthList[j];
                        builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
                        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;//垂直居中对齐
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐
                        builder.Write(nameList.Rows[i][j].ToString());
                    }
                   
                    builder.EndRow();
                    
                }
                doc.Save(saveDocFile);
                doc.Range.Bookmarks["table"].Text = "";    // 清掉标示 
                //doc.Save(saveDocFile);
                
                    System.Diagnostics.Process.Start(saveDocFile);
                
            }
            catch (Exception ex)
            {
                throw ex;
                
            }
        }
    }
}