using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;

namespace BillyCasperGAssignment.Models
{
    // You may need to install the Microsoft.AspNetCore.Http.Abstractions package into your project
    public class Costumer
    {

        public int ID { get; set; }

        public DateTime CreatedOn { get; set; }

        public DateTime ModifiedOn { get; set; }

        public string Costumer_LastName { get; set; }

        public string Costumer_FirstName { get; set; }

        public string AddressLine1 { get; set; }

        public string Costumer_City { get; set; }

        public string Costumer_State { get; set; }

        public string Costumer_zip { get; set; }

        public string Costumer_Homephone { get; set; }

        public string Costumer_InternetEmail { get; set; }
    }
}
