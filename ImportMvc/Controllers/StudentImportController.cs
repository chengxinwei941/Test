using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
namespace ImportMvc.Controllers
{
    public class StudentImportController : Controller
    {
        // GET: StudentImport
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public void Import()
        {
            #region 导入到DataTable
            //接收上传的文件
            HttpPostedFileBase fileBase = Request.Files["fileExcel"];
            //是否上传文件
            if (fileBase != null)
            {
                if (Path.GetExtension(fileBase.FileName).ToLower().Equals(".xlsx"))
                {
                    //获取文件流 
                    Stream stream = fileBase.InputStream;
                    //数据流转工作簿
                    IWorkbook workbook = new XSSFWorkbook(stream);
                    //获取sheet
                    ISheet sheet = workbook.GetSheetAt(0);
                    //定义数据表
                    DataTable dt = new DataTable();
                    //获取表头
                    IRow row = sheet.GetRow(1);
                    foreach (ICell item in row.Cells)
                    {
                        //表头
                        dt.Columns.Add(item.StringCellValue);
                    }

                    for (int i = 2; i < sheet.LastRowNum + 1; i++)
                    {
                        //创建行
                        DataRow dr = dt.NewRow();
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            ICell cell = sheet.GetRow(i).Cells[i];
                            SetRowValue(sheet, i, dr, j);
                        }
                        dt.Rows.Add(dr);
                    }

                    InsetData(dt);
                }

            }
            #endregion
        }

        private static void SetRowValue(ISheet sheet, int i, DataRow dr, int j)
        {
            switch (sheet.GetRow(i).Cells[j].CellType)
            {
                case CellType.String: dr[j] = sheet.GetRow(i).Cells[j].StringCellValue; break;
                case CellType.Numeric: dr[j] = sheet.GetRow(i).Cells[j].NumericCellValue; break;
                case CellType.Boolean: dr[j] = sheet.GetRow(i).Cells[j].BooleanCellValue; break;
                default:
                    dr[j] = sheet.GetRow(i).Cells[j].ErrorCellValue;
                    break;
            }
        }

        public void InsetData(DataTable dt)
        {
            string strConn = ConfigurationManager.ConnectionStrings["sqlconn"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(strConn))
            {
                conn.Open();
                SqlCommand cmd = conn.CreateCommand();
                using (SqlTransaction tran = conn.BeginTransaction())
                {
                    cmd.Transaction = tran;
                    try
                    {
                        foreach (DataRow item in dt.Rows)
                        {
                            cmd.CommandText = $"insert into StudentInfo values(";
                            for (int i = 0; i < dt.Columns.Count - 1; i++)
                            {
                                if (i == 3)
                                {
                                    cmd.CommandText += item[i].ToString() == "男" ? "1," : "0,";
                                    continue;
                                }
                                if (i == 5)
                                {
                                    cmd.CommandText += $"(select classId from ClassInfo where ClassName='{item[i]}'),";
                                    continue;
                                }
                                cmd.CommandText += $"'{item[i]}',";
                            }
                            cmd.CommandText = cmd.CommandText.TrimEnd(',');
                            cmd.CommandText += ")";
                            cmd.ExecuteNonQuery();
                           
                        }
                        tran.Commit();
                    }
                    catch (Exception)
                    {
                        tran.Rollback();
                        throw;
                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                        cmd.Dispose();
                        tran.Dispose();
                    }
                }

            }
        }


    }
}