using Dapper;
using Excel_Database.Domain.Interface;
using Excel_Databse.Repository;
using Excel_Databse.Service;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Npgsql;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace Excel_Database.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HomeController : Controller
    {
        private readonly IDBService _service;
        
        public HomeController(IDBService service)
        {
            _service = service;
        }

        [HttpPost("Excelimport")]
        public async Task<Object> ImportExcelData(IFormFile file,string tablename)
        {
            var res=  await _service.ImportExcelData(file, tablename);
            if (res.ErrorCode == "00")
            {
                return Ok("Data Imported successfully To Database !");
            }
            else
                return res.ErrorMessage;
        }
    }
}
