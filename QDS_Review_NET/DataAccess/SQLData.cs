using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QDS_Review_NET;
using QDS_Review_NET.Models;

namespace QDS_Review_NET.DataAccess
{
    public class SQLData : Helpers
    {
        public List<EmployeesReportsModel> Get_Emp_Reports(string brh, string name)
        {
            List<EmployeesReportsModel> lst = new List<EmployeesReportsModel>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Emails_ByBrh_ByRptName";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 2, ParameterDirection.Input, brh);
                AddParamToSQLCmd(cmd, "@name", SqlDbType.VarChar, 25, ParameterDirection.Input, name);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        EmployeesReportsModel r = new EmployeesReportsModel();

                        r.rptID = (int)rdr["RptID"];
                        r.email = (string)rdr["EmpEmail"];
                        r.temppath = (string)rdr["RptTempPath"];
                        r.rootpath = (string)rdr["RptRootPath"];
                        r.filename = (string)rdr["RptFileName"];
                        r.fullpath = (string)rdr["RptFullPath"];                       

                        lst.Add(r);
                    }
                }
            }

            return lst;
        }

        public List<string> Get_PWC_ByBrh(string brh)
        {
            List<string> pwcList = new List<string>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            try
            {
                using (conn)
                {
                    conn.Open();

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "PWC_LKU_proc_Active_ByBrh";
                    cmd.Connection = conn;

                    AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 3, ParameterDirection.Input, brh);

                    rdr = cmd.ExecuteReader();

                    using (rdr)
                    {
                        while (rdr.Read())
                        {
                            string s;

                            s = (string)rdr["PWC"];

                            pwcList.Add(s);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return pwcList;
        }
    }
}
