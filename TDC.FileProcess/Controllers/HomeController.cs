using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Web;
using System.Web.Mvc;
using TDC.FileProcess.Dtos;
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
                            //conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                            conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                            break;

                        case ".xlsx": //Excel 07 and above.
                            //conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
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

                            List<Employees> arrRows = new List<Employees>();
                            List<Employees> arrColumns = new List<Employees>();

                            var rawNumber = dt.Rows.Count;
                            var colNumber = dt.Columns.Count;
                            try
                            {
                                for (int i = 0; i < rawNumber; i++)
                                {
                                    var countCol = dt.Rows[i].ItemArray;
                                    var code = dt.Rows[i].ItemArray[0];
                                    var fullName = dt.Rows[i].ItemArray[1];
                                    var department = dt.Rows[i].ItemArray[2];
                                    DateTime dateWorking = (DateTime)dt.Rows[i].ItemArray[3];

                                    sqlBulkCopy.ColumnMappings.Add("Code", "Code");
                                    sqlBulkCopy.ColumnMappings.Add("FullName", "FullName");
                                    sqlBulkCopy.ColumnMappings.Add("Department", "Department");
                                    sqlBulkCopy.ColumnMappings.Add("DateWorking", "DateWorking");

                                    bool stepIn = false;
                                    bool stepOut = false;

                                    for (int j = 4; j < countCol.Length; j++)
                                    {

                                        if (!dt.Rows[i].ItemArray[j.ToInt()].IsNullOrEmptyOrWhileSpace())
                                        {
                                            DateTime checkIn = (DateTime)dt.Rows[i].ItemArray[j.ToInt()];
                                            if (!checkIn.IsNullOrEmptyOrWhileSpace())
                                            {
                                                if (checkIn.Hour < 13 && stepIn != true)
                                                {
                                                    sqlBulkCopy.ColumnMappings.Add("CheckIn", "CheckIn");
                                                    stepIn = true;
                                                }
                                                if (checkIn.Hour > 13 && stepOut != true)
                                                {
                                                    sqlBulkCopy.ColumnMappings.Add("CheckIn", "CheckOut");
                                                    stepOut = true;
                                                }
                                            }
                                        }
                                    }
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(dt);
                                    con.Close();
                                }
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


    }
}