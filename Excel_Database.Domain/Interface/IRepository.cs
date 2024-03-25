using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Database.Domain.Interface
{
    public  interface IRepository
    {
        public Task<object> ImportExcel(IFormFile file, string tablename);
    }
}
