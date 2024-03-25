using Dapper;
using Microsoft.AspNetCore.Mvc;
using Npgsql;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;

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

                try
                {
                    connection.Open();

                    // Check if the connection is open
                    if (connection.State == ConnectionState.Open)
                    {
                        using var command1 = new NpgsqlCommand("CREATE TABLE \"Testtable\" (\"Id\" SERIAL PRIMARY KEY, \"Name\" VARCHAR(100), \"Age\" VARCHAR(100))", connection);
                        command1.ExecuteNonQuery();
                    }
                    else
                    {
                        Console.WriteLine("Connection failed to open.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }

                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null)
                    {
                        continue;
                    }

                    string column1Value = row.GetCell(0)?.ToString();
                        string column2Value = row.GetCell(1)?.ToString();
                        string column3Value = row.GetCell(2)?.ToString();
                        string column4Value = row.GetCell(3)?.ToString();
                        string column5Value = row.GetCell(4)?.ToString();
                        string column6Value = row.GetCell(5)?.ToString();
                        string column7Value = row.GetCell(6)?.ToString();
                        string column8Value = row.GetCell(7)?.ToString();
                        string column9Value = row.GetCell(8)?.ToString();
                        string column10Value = row.GetCell(9)?.ToString();
                        string column11Value = row.GetCell(10)?.ToString();
                        
                    
                    if (tablename == "SR")
                    {
                        using var command = new NpgsqlCommand(
                            "INSERT INTO \"SR\" (\"SRPcode\",\"SRNameEng\",\"SRNameMMR\") VALUES (@column1, @column2,@column3)",
                            connection);
                        command.Parameters.AddWithValue("column1", column1Value);
                        command.Parameters.AddWithValue("column2", column2Value);
                        command.Parameters.AddWithValue("column3", column3Value);
                        command.ExecuteNonQuery();
                                           }
                   
                    if (tablename == "District")
                    {
                        using var command2 = new NpgsqlCommand("INSERT INTO \"District\" (\"SR\",\"SRNameEng\",\"DistrictPcode\",\"DistrictNameEng\",\"DistrictNameMMR\" )" +
                            " VALUES (@column1,@column2,@column3,@column4,@column5)", connection);
                        command2.Parameters.AddWithValue("column1", column1Value);
                        command2.Parameters.AddWithValue("column2", column2Value);
                        command2.Parameters.AddWithValue("column3", column3Value);
                        command2.Parameters.AddWithValue("column4", column4Value);
                        command2.Parameters.AddWithValue("column5", column5Value);
                        command2.ExecuteNonQuery();
                  
                     

                    }
                    if (tablename == "Township")
                    {
                        using var command = new NpgsqlCommand("INSERT INTO \"Township\" (\"SRPcode\",\"SRNameEng\",\"DistrictPcode\",\"DistrictNameEng\",\"TspPcode\",\"TownshipNameEng\",\"TownshipNameMMR\")  VALUES (@column1,@column2,@column3,@column4,@column5,@column6,@column7)", connection);
                        command.Parameters.AddWithValue("column1", column1Value);
                        command.Parameters.AddWithValue("column2", column2Value);
                        command.Parameters.AddWithValue("column", column3Value);
                        command.Parameters.AddWithValue("column4", column4Value);
                        command.Parameters.AddWithValue("column5", column5Value);
                        command.Parameters.AddWithValue("column6", column6Value);
                        command.Parameters.AddWithValue("column7", column7Value);                      
                        command.ExecuteNonQuery();
          
                    }

                    if (tablename == "Town")
                    {
                         
                        using var command = new NpgsqlCommand("INSERT INTO \"Town\" (\"SRPcode\",\"SRNameEng\",\"DistrictPcode\",\"DistrictNameEng\",\"TspPcode\",\"TownshipNameEng\",\"TownPcode\",\"TownNameEng\",\"TownNameMMR\") " +
                          "VALUES (@column1,@column2,@column3,@column4,@column5,@column6,@column7,@column8,@column9)", connection);
                        command.Parameters.AddWithValue("column1", column1Value);
                        command.Parameters.AddWithValue("column2", column2Value);
                        command.Parameters.AddWithValue("column3", column3Value);
                        command.Parameters.AddWithValue("column4", column4Value);
                        command.Parameters.AddWithValue("column5", column5Value);
                        command.Parameters.AddWithValue("column6", column6Value);
                        command.Parameters.AddWithValue("column7", column7Value);
                        command.Parameters.AddWithValue("column8", column8Value);
                        command.Parameters.AddWithValue("column9", column9Value);
                        command.ExecuteNonQuery();
                    }
                    if(tablename == "Ward")
                    {
                        using var command = new NpgsqlCommand("INSERT INTO \"Ward\" (\"SRPcode\",\"SRNameEng\",\"DistrictPcode\",\"DistrictNameEng\",\"TspPcode\",\"TownshipNameEng\",\"TownPcode\",\"Town\",\"WardPcode\",\"WardNameEng\",\"WardNameMMR\")    VALUES (@column1,@column2,@column3,@column4,@column5,@column6,@column7,@column8,@column9,@column10,@column11)    ", connection);
                        command.Parameters.AddWithValue("column1", column1Value);
                        command.Parameters.AddWithValue("column2", column2Value);
                        command.Parameters.AddWithValue("column3", column3Value);
                        command.Parameters.AddWithValue("column4", column4Value);
                        command.Parameters.AddWithValue("column5", column5Value);
                        command.Parameters.AddWithValue("column6", column6Value);
                        command.Parameters.AddWithValue("column7", column7Value);
                        command.Parameters.AddWithValue("column8", column8Value);
                        command.Parameters.AddWithValue("column9", column9Value);
                        command.Parameters.AddWithValue("column10", column10Value);
                        command.Parameters.AddWithValue("column11", column11Value);
 
                        command.ExecuteNonQuery();
                      
                    }
                   
             
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
