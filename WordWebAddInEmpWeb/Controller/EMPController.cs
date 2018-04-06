using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WordWebAddInEmpWeb.Model;

namespace WordWebAddInEmpWeb
{
    public class EMPController : ApiController
    {
        EmployeeDataAccess objdep = new EmployeeDataAccess();
        // GET: api/<controller>


        [HttpGet()]
 //[Route("api/EMP/Index")]
        public IEnumerable<Employee> EMP()
        {
            return objdep.GetAllEmployee();
        }
    }
}