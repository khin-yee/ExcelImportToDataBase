using Excel_Database.Domain.Interface;
using Excel_Database.Domain.Model;
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

        public async Task<Response> ImportExcelData(IFormFile filestring ,string tablename)
        {
            Response res =  _repo.ImportExcelData(filestring, tablename);
            return res;            
        }
    }
}