using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace testexcel.Controllers
{
    public class ExcelController : Controller
    {
        [Obsolete]
        private IHostingEnvironment _ev;
        private IConfiguration _config;

        [Obsolete]
        public ExcelController(IHostingEnvironment ev, IConfiguration config)
        {
            _ev = ev;
            _config = config;
        }

        public IActionResult Index()
        {
            return View();
        }

        
        [HttpPost]
        [Obsolete]
        public IActionResult Index(IFormFile file)
        {
            if (file != null)
            {
                string path = Path.Combine(_ev.WebRootPath, "files");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string fileName = Path.GetFileName(file.FileName);
                string filePath = Path.Combine(path, fileName);
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }


                string connString = this._config.GetConnectionString("ExcelConnectionString");
                DataTable dt = new DataTable();
                connString = string.Format(connString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(connString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;
                            connExcel.Open();
                            DataTable dtExcel;
                            dtExcel = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcel.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }
                }

                connString = this._config.GetConnectionString("ConnString");
                using (SqlConnection conn = new SqlConnection(connString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(conn))
                    {
                        sqlBulkCopy.DestinationTableName = "[User]"; 
                        // sqlBulkCopy.ColumnMappings.Add("Id", "CustomerId");
                        // sqlBulkCopy.ColumnMappings.Add("Name", "Name");
                        // sqlBulkCopy.ColumnMappings.Add("Country", "Country");

                        conn.Open();
                        sqlBulkCopy.WriteToServer(dt);
                        conn.Close();
                    }
                }
            }

            return View();
        }
    }
}