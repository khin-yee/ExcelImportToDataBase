using Excel_Database.Domain.Interface;
using Excel_Database.Domain.Model;
using Microsoft.AspNetCore.Http;
using Npgsql;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;

namespace Excel_Databse.Repository
{
    public class DBExcelRepository:IRepository
    {
        protected readonly IDbConnection _connection;
        public DBExcelRepository(IDbConnection connection)
        {
            _connection = connection;
        }
        public  Response  ImportExcelData(IFormFile file, string tablename)
        {
            Response res = new Response();
            try
            {

                if (file == null || file.Length == 0)
                {
                    res.ErrorMessage = "This file is empty";
                    res.ErrorCode = "01";
                }
                using var stream = file.OpenReadStream();
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);
                if(sheet.ActiveCell == null)
                {
                    res.ErrorMessage = "This file has no data";
                    res.ErrorCode = "01";
                    return res;
                }
                _connection.Open();

                // Check if the _connection is open
                if (_connection.State == ConnectionState.Open)
                {
                    string createTableQuery = $"CREATE TABLE \"{tablename}\" (\"Id\" SERIAL PRIMARY KEY";
                    string insertquery = $"Insert into \"{tablename}\"(";

                    IRow row = sheet.GetRow(0);
                    for (int i = 0; i < row.Cells.Count; i++)
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
                     
                    using var command1 = new NpgsqlCommand(createTableQuery, (NpgsqlConnection?)_connection);
                    command1.ExecuteNonQuery();
                    using var command2 = new NpgsqlCommand(insertquery, (NpgsqlConnection?)_connection);
                    for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
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
                    using var insertcommand = new NpgsqlCommand(insertquery, (NpgsqlConnection?)_connection);
                    insertcommand.ExecuteNonQuery();
                    return res;
                }
  
            }
            catch (Exception ex)
            {
                res.ErrorMessage = ex.Message;
                res.ErrorCode = "01";
                return res;
            }
            return res;
        }

    }
}