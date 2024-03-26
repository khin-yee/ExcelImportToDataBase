using Dapper;
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
        private readonly  string connectionString = "Host=localhost;Port= 5432;User ID=postgres;Password=762000;Database=TestDatabase";

        [HttpPost("Excelimport")]
        public IActionResult ImportExcelData(IFormFile file,string tablename)
        {
            try
            {
                
                if (file == null || file.Length == 0)
                {
                    return BadRequest("No file uploaded.");
                }
                using var stream = file.OpenReadStream();
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);

                using var connection = new NpgsqlConnection(connectionString);

                    connection.Open();

                    // Check if the connection is open
                    if (connection.State == ConnectionState.Open)
                    {
                        string createTableQuery = $"CREATE TABLE \"{tablename}\" (\"Id\" SERIAL PRIMARY KEY";
                        string insertquery = $"Insert into \"{tablename}\"(";

                        IRow row = sheet.GetRow(0);
                        for (int i= 0; i < row.Cells.Count; i++)
                        {
                            string columnValue = row.GetCell(i)?.ToString();

                            createTableQuery += $", \"{columnValue}\"  VARCHAR(100)";
                            if (i == 0)
                            {
                                insertquery += $"\"{columnValue}\"";
                            }
                            else
                            {
                                insertquery += $",\"{columnValue}\"";
                            }
                        }

                        createTableQuery += ")";
                        insertquery += ") VALUES ";
                        using var command1 = new NpgsqlCommand(createTableQuery, connection);
                        command1.ExecuteNonQuery();
                        using var command2 = new NpgsqlCommand(insertquery, connection);
                        for (int rowIndex = 1;rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            IRow rows = sheet.GetRow(rowIndex);
                            insertquery += "(";
                            for (int i = 0; i < row.Cells.Count; i++)
                            {

                                string columnValue = rows.GetCell(i)?.ToString();
                                if (i == 0)
                                {
                                    insertquery += $"'{columnValue}'";
                                }
                                else
                                {
                                    insertquery += $",'{columnValue}'";

                                }
                                

                            }
                            if (rowIndex == sheet.LastRowNum)
                            {
                                insertquery += ")";
                            }
                            else
                            {
                                insertquery += "),";
                            }
                          
                        }
                        using var insertcommand = new NpgsqlCommand(insertquery, connection);
                        insertcommand.ExecuteNonQuery();


                    }
                    else
                    {
                                       return StatusCode(500, $"An error occurred:");

                    }
             
                return Ok("Data imported successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred: {ex.Message}");
            }
        }
    }
}
