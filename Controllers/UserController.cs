using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using testexcel.Models;

namespace testexcel.Controllers
{
    public class UserController : Controller
    {
        [HttpGet]
        public IActionResult Index(List<User> users = null)
        {
            users = users == null ? new List<User>() : users;
            return View(users);
        }

        [HttpPost]
        [Obsolete]
        public IActionResult Index(IFormFile file, [FromServices] IHostingEnvironment hostingEnvironment)
        {
            string fileName = $"{hostingEnvironment.WebRootPath}/files/{file.FileName}";
            using (FileStream fileStream = System.IO.File.Create(fileName))
            {
                file.CopyTo(fileStream);
                fileStream.Flush();
            }

            var user = this.GetUserList(file.FileName);
            return Index(user);
        }

        private List<User> GetUserList(string fName)
        {
            List<User> users = new List<User>();
            var fileName = $"{Directory.GetCurrentDirectory()}{@"/wwwroot/files"}" + "/" + fName;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        users.Add(new User()
                        {
                            Name = reader.GetValue(0).ToString(),
                            Age = reader.GetValue(1).ToString(),
                            FirstName = reader.GetValue(2).ToString()
                        });
                    }
                }
            }
            return users;
        }
    }
}
