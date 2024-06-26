﻿using Excel_Database.Domain.Model;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Database.Domain.Interface
{
    public interface IDBService
    {
        public Task<Response> ImportExcelData(IFormFile file, string tablename);

    }
}
