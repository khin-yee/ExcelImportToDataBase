using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Database.Domain.Model
{
    public  class Response
    {
        public string ErrorCode { get; set; } = "00";

        public string ErrorMessage { get; set; } = "No Error";
    }
}
