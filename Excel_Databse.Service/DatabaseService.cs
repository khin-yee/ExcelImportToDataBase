using Excel_Database.Domain.Interface;
using Excel_Databse.Repository;
using Microsoft.AspNetCore.Http;
using System.Reflection.Metadata.Ecma335;

namespace Excel_Databse.Service
{
    public class DatabaseService:IDBService
    {

        private readonly IRepository _repo;

        public DatabaseService(IRepository repo)
        {
            _repo = repo;
        }

        public async Task ImportExcelData(IFormFile filestring ,string tablename)
        {
            await _repo.ImportExcelData(filestring, tablename);
        }
    }
}