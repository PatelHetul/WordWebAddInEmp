using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordWebAddInEmpWeb.Model
{
    public class Employee
    {
        public int Employee_Id { get; set; }
      
        public string Employee_Name { get; set; }

        public string Department { get; set; }

        public Nullable<System.DateTime> JoiningDate { get; set; }

        public string Address { get; set; }

        public string Email { get; set; }

        public string MobileNo { get; set; }
       
    }
}