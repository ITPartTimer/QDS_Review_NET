using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Data.OleDb;
using System.Text;
using System.Threading.Tasks;
using QDS_Review_NET.Models;

namespace QDS_Review_NET.DataAccess
{
    public class ExcelExport
    {
        public void WriteQDSNotAppv(List<QDSDataModel> lstData, string fullPath)
        {
            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    foreach (QDSDataModel m in lstData)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [NotAppv$] (WHS,TAG,INV_TYP,INV_STS,REF_PFX,REF_NO,QDS_NO,QDS_DT) Values(@whs,@tag,@typ,@sts,@pfx,@ref,@qds,@dt)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@whs", m.whs),
                                    new OleDbParameter("@tag", m.tag),
                                    new OleDbParameter("@typ", m.typ),
                                    new OleDbParameter("@sts", m.sts),
                                    new OleDbParameter("@pfx", m.ref_pfx),
                                    new OleDbParameter("@ref", m.ref_no),
                                    new OleDbParameter("@qds", m.qds),
                                    new OleDbParameter("@dt", m.dt),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException)
                {
                    throw;
                }
                catch(Exception)
                {
                    throw;
                }
                finally
                {
                    eConn.Close();
                    eConn.Dispose();
                }
            }
        }
    }
    
}
