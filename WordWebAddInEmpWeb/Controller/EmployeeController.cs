using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace WordWebAddInEmpWeb.Controller
{
    public class EmployeeController : ApiController
    {

        public class Employee
        {
            public int emp_id { get; set; }
            public string employee_Name { get; set; }
            public string department_Name { get; set; }
            public string joiningDate { get; set; }
            public string address { get; set; }
            public string email { get; set; }
            public string mobileNo { get; set; }
        }
        public class EditEmp
        {
            public string Colname { get; set; }
            public string Value { get; set; }
        }

        [HttpGet()]
        public IEnumerable<Employee> employeeList(string empid)
        {
            int emp_id = 0;
            int.TryParse(empid.ToString(), out emp_id);

            int query = 1;
            if (emp_id > 0)
            {
                query = 3;
            }
            List<Employee> lstemployee = new List<Employee>();
            using (SqlConnection con = new SqlConnection("Data Source = (localdb)\\mssqllocaldb; Initial Catalog = EmployeeManagement; Integrated Security = True"))
            {
                SqlCommand cmd = new SqlCommand("ManageEmployee", con);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Query", query);
                cmd.Parameters.AddWithValue("@empid", emp_id);
                cmd.Parameters.AddWithValue("@name", "");
                cmd.Parameters.AddWithValue("@depart", "");
                cmd.Parameters.AddWithValue("@joining", "");
                cmd.Parameters.AddWithValue("@address", "");
                cmd.Parameters.AddWithValue("@Email", "");
                cmd.Parameters.AddWithValue("@mobile", "");
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    if (emp_id == 0)
                    {
                        Employee emp = new Employee();
                        emp.emp_id = int.Parse(rdr["Employee_Id"].ToString());
                        emp.employee_Name = rdr["Employee_Name"].ToString();
                        emp.joiningDate = rdr["JoiningDate"].ToString();
                        emp.department_Name = rdr["Department_Name"].ToString();
                        emp.email = rdr["Email"].ToString();
                        emp.address = rdr["Address"].ToString();
                        emp.mobileNo = rdr["MobileNo"].ToString();

                        lstemployee.Add(emp);
                    }
                    else
                    {
                        if (int.Parse(rdr["Employee_Id"].ToString()) == emp_id)
                        {
                            Employee emp = new Employee();

                            emp.employee_Name = "Employee Name";
                            emp.email = rdr["Employee_Name"].ToString();
                            lstemployee.Add(emp);

                            emp = new Employee();
                            emp.employee_Name = "Joining Date";
                            emp.email = rdr["JoiningDate"].ToString();
                            lstemployee.Add(emp);

                            emp = new Employee();
                            emp.employee_Name = "Department Name";
                            emp.email = rdr["Department_Name"].ToString();
                            lstemployee.Add(emp);

                            emp = new Employee();
                            emp.employee_Name = "Email";
                            emp.email = rdr["Email"].ToString();
                            lstemployee.Add(emp);

                            emp = new Employee();
                            emp.employee_Name = "Address";
                            emp.email = rdr["Address"].ToString();
                            lstemployee.Add(emp);

                            emp = new Employee();
                            emp.employee_Name = "Mobile No";
                            emp.email = rdr["MobileNo"].ToString();
                            lstemployee.Add(emp);
                        }
                    }
                }
                con.Close();
            }

            return lstemployee;
        }

        [HttpGet()]
        [Route("api/Employee/empid")]
        public int EditEmployee(string empid, string name, string date, string depart, string emails, string add, string mobileno)
        {
            try
            {
                int emp_id = 0;
                int.TryParse(empid.ToString(), out emp_id);
                int retval = 0;
                using (SqlConnection con = new SqlConnection("Data Source = (localdb)\\mssqllocaldb; Initial Catalog = EmployeeManagement; Integrated Security = True"))
                {
                    SqlCommand cmd = new SqlCommand("ManageEmployee", con);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Query", 2);
                    cmd.Parameters.AddWithValue("@empid", emp_id);
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@depart", depart);
                    cmd.Parameters.AddWithValue("@joining", date);
                    cmd.Parameters.AddWithValue("@address", add);
                    cmd.Parameters.AddWithValue("@Email", emails);
                    cmd.Parameters.AddWithValue("@mobile", mobileno);

                    con.Open();
                    retval = cmd.ExecuteNonQuery();
                    con.Close();
                }
                return retval;
            }
            catch (Exception aa)
            {
                return  0;
            }

        }

    }
}