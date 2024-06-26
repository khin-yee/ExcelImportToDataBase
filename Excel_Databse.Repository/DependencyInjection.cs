﻿using Excel_Database.Domain.Interface;
using Microsoft.Extensions.DependencyInjection;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Databse.Repository
{
    public static  class DependencyInjection
    {
        public  static IServiceCollection AddDapperService(this IServiceCollection service)
        {
            service.AddScoped<IDbConnection>((sp) => new NpgsqlConnection("Host=localhost;Port= 5432;User ID=postgres;Password=762000;Database=AccountPjDB"));
            service.AddScoped<IRepository, DBExcelRepository>();
            return service;
        }
    }
}
