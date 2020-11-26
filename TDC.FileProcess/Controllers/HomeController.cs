using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Web;
using System.Web.Mvc;
using TDC.FileProcess.Ultilities;

namespace TDC.FileProcess.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult uploadData(HttpPostedFileBase postedFile)
        {
            string filePath = string.Empty;
           
            if (postedFile != null)
            {
                try
                {
                    string extension = Path.GetExtension(postedFile.FileName);//Lấy ra đuôi của file

                    if (extension == ".xls" || extension == ".xlsx")
                    {
                        string filename = string.Empty;
                        if (extension == ".xls")
                        {
                            filename = Request.Files["postedFile"].FileName.Substring(0, Request.Files["postedFile"].FileName.Length - 4);
                            filename = filename + "_" + "_" + DateTime.Now.ToString().Replace(":", "").Replace("/", "").Replace(" ", "") + ".xls";
                        }
                        else
                        {
                            filename = Request.Files["postedFile"].FileName.Substring(0, Request.Files["postedFile"].FileName.Length - 5);
                            filename = filename + "_" + DateTime.Now.ToString().Replace(":", "").Replace("/", "").Replace(" ", "") + ".xlsx";
                        }

                        string path = Server.MapPath("~/upload/") + filename;
                        //Nếu chưa tồn tại thư mục thì tạo thư mục
                        if (!Directory.Exists(Server.MapPath("~/upload/")))
                        {
                            Directory.CreateDirectory(path);
                        }
                        filePath = path;
                        //Nếu đã tồn tại file thì xóa file
                        if (System.IO.File.Exists(path))
                        {
                            System.IO.File.Delete(path);
                        }
                    }
                    postedFile.SaveAs(filePath);
                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".xls": //Excel 97-03.
                            conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                            break;

                        case ".xlsx": //Excel 07 and above.
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                            break;
                        default:
                            Console.WriteLine("Khác");
                            break;
                    }

                    DataTable dt = new DataTable();
                    conString = string.Format(conString, filePath);

                    using (OleDbConnection connExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = new OleDbCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = connExcel;

                                //Get the name of First Sheet.
                                connExcel.Open();
                                DataTable dtExcelSchema;
                                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                connExcel.Close();

                                //Read Data from First Sheet.
                                connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                connExcel.Close();
                            }
                        }
                    }

                    conString = ConfigurationManager.ConnectionStrings["DBContext"].ConnectionString;

                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name.
                            sqlBulkCopy.DestinationTableName = "dbo.Files";
                            var rawNumber = dt.Rows.Count;
                            var colNumber = dt.Columns.Count;
                            try
                            {
                                int startrowMain = 3;

                                 var templateFilePath = Server.MapPath("/Content/TemplateExcel/Report.xlsx");
                                    FileInfo m_FileInfo = new FileInfo(templateFilePath);
                                byte[] dataByte;
                                using (var excelfilecontent = new ExcelPackage(m_FileInfo))
                                {
                                    for (int i = 0; i < rawNumber; i++)
                                    {
                                        var code       = dt.Rows[i].ItemArray[0].ToString();
                                        var fullName   = dt.Rows[i].ItemArray[1].ToString();
                                        var department = dt.Rows[i].ItemArray[2].ToString();
                                        var dayWorking = dt.Rows[i].ItemArray[3].ToString("dd/MM/yyyy");

                                        int isNumData = 0;
                                        int tmp = 1;
                                        //Đếm tổng số lần user quét trên từng dòng một.
                                        for (int j = 4; j < colNumber; j++)
                                        {
                                            if (!dt.Rows[i].ItemArray[j.ToInt()].IsNullOrEmptyOrWhileSpace())
                                            {
                                                isNumData++;
                                            }
                                        }

                                        var worksheet1 = excelfilecontent.Workbook.Worksheets[0];
                                        //Bắt đầu insert từ dòng 17 trong template
                                        Insert_RichText(ref worksheet1, startrowMain, 1, code.IsNullOrEmptyOrWhileSpace()? "Day off" : code, false);
                                        Insert_RichText(ref worksheet1, startrowMain, 2, fullName.IsNullOrEmptyOrWhileSpace() ? "Day off" : fullName, false);
                                        Insert_RichText(ref worksheet1, startrowMain, 3, department.IsNullOrEmptyOrWhileSpace() ? "Day off" : department, false);
                                        Insert_RichText(ref worksheet1, startrowMain, 4, dayWorking.IsNullOrEmptyOrWhileSpace() ? "Day off" : dayWorking, false);

                                        //Lặp qua tổng số lần, và lấy lần đầu tiên và lần cuối cùng
                                        for (int item = 4; item < colNumber; item++)
                                        {
                                            if (tmp == 1)
                                            {
                                                Insert_RichText(ref worksheet1, startrowMain, 5, dt.Rows[i].ItemArray[item].IsNullOrEmptyOrWhileSpace() ? "Day off" : dt.Rows[i].ItemArray[item].ToString("hh:mm tt"), false);
                                            }
                                            if (tmp == isNumData)
                                            {
                                                Insert_RichText(ref worksheet1, startrowMain, 6, dt.Rows[i].ItemArray[item].IsNullOrEmptyOrWhileSpace() ? "Day off" : dt.Rows[i].ItemArray[item].ToString("hh:mm tt"), false);
                                            }
                                            tmp++;
                                        }
                                        startrowMain ++;
                                    }

                                    dataByte = excelfilecontent.GetAsByteArray();
                                }
                                return File(dataByte, "application/xlsx", "CHAMCONG-" + DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss") + ".xlsx");
                            }
                            catch (Exception ex)
                            {
                                ex.ToString();
                                throw;
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
            return Redirect("~/Home");
        }

        public void Insert_RichText(ref ExcelWorksheet worksheet, int IndexRow, int IndexColumn, string Value, bool isBold, bool isRed = false)
        {
            worksheet.Cells[IndexRow, IndexColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[IndexRow, IndexColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ExcelRichText RichText1 = worksheet.Cells[IndexRow, IndexColumn].RichText.Add(Value);
            if (isRed)
            {
                RichText1.Bold = false;
                RichText1.Italic = false;
                RichText1.Color = Color.Red;
                RichText1.FontName = "Arial Narrow";
            }
            else
            {
                RichText1.Color = Color.Black;
            }

            worksheet.Cells[IndexRow, IndexColumn].Style.WrapText = true;
            worksheet.Cells[IndexRow, IndexColumn].IsRichText = true;
            worksheet.Row(IndexRow).Height = 20;

            //Auto height row
            worksheet.Row(IndexRow).CustomHeight = false;

            RichText1.Bold = isBold;
        }

        public void Insert_Table_Column(ref ExcelWorksheet worksheet, int IndexRow, int IndexColumn, object Value, bool isCenter = false, bool isBold = false)
        {
            worksheet.Cells[IndexRow, IndexColumn].Value = Value;
            worksheet.Cells[IndexRow, IndexColumn].Style.Font.Bold = isBold;
            worksheet.Cells[IndexRow, IndexColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[IndexRow, IndexColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            if (isCenter)
                worksheet.Cells[IndexRow, IndexColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[IndexRow, IndexColumn].Style.Font.Color.SetColor(System.Drawing.Color.Black);
            worksheet.Cells[IndexRow, IndexColumn].Style.WrapText = true;
            worksheet.Cells[IndexRow, IndexColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
        public void Set_Cell_Border(ref ExcelWorksheet worksheet, int IndexRow, int IndexColumn)
        {
            worksheet.Cells[IndexRow, IndexColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
    }
}