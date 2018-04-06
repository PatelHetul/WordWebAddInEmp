using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Services;

namespace WordWebAddInEmpWeb.Model
{
    public class EmployeeDataAccess
    {
        static string connectionString = "Data Source = (localdb)\\mssqllocaldb; Initial Catalog = EmployeeManagement; Integrated Security = True";
      
        public  IEnumerable<Employee> GetAllEmployee()
        {
            try
            {
                List<Employee> lstemployee = new List<Employee>();

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand("ManageEmployee", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Query", 1);
                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();

                    while (rdr.Read())
                    {
                        Employee emp = new Employee();

                        emp.Employee_Id = Convert.ToInt32(rdr["Employee_Id"]);
                        emp.Employee_Name = rdr["Employee_Name"].ToString();
                        emp.JoiningDate = DateTime.Parse(rdr["JoiningDate"].ToString());
                        emp.Department = rdr["Department_Name"].ToString();
                        emp.Email = rdr["Email"].ToString();
                        emp.Address = rdr["Address"].ToString();
                        emp.MobileNo = rdr["MobileNo"].ToString();

                        lstemployee.Add(emp);
                    }
                    con.Close();
                }
                return lstemployee;
            }
            catch
            {
                throw;
            }
        }

        [WebMethod]
        public static string list()
        {
            return "hello world";
        }
    }
}