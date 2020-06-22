using System;
using System.Data.Odbc;
using System.Collections.Generic;
using System.Text;
using QDS_Review_NET;
using QDS_Review_NET.Models;

namespace QDS_Review_NET.DataAccess
{
    public class ODBCData : Helpers
    {
        public List<QDSDataModel> Get_QDSNotAppv()
        {
            List<QDSDataModel> lstQDSDataModel = new List<QDSDataModel>();

            OdbcConnection conn = new OdbcConnection(ODBCDataConnString);

            try
            {
                conn.Open();

                // Try to split with verbatim literal
                OdbcCommand cmd = conn.CreateCommand();

                cmd.CommandText = @"select prd_whs,prd_tag_no,prd_invt_typ,prd_invt_sts,pcr_crtd_ref_pfx,pcr_crtd_ref_no,
                                    NVL(qds_qds_ctl_no,'0') as qds_qds_ctl_no,NVL(qds_crtd_dt, '0') as qds_crtd_dt
                                    from intprd_rec prd
                                    inner join intpcr_rec pcr on prd.prd_itm_ctl_no = pcr.pcr_itm_ctl_no
                                    left outer join mchqds_rec qds on pcr.pcr_qds_ctl_no = qds.qds_qds_ctl_no
                                    where (prd_invt_typ  = 'M' or prd_invt_typ = 'F')
                                    and prd_invt_sts = 'S'
                                    and qds_apvd_lgn_id is null
                                    and prd_ownr = 'O'";

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        QDSDataModel b = new QDSDataModel();

                        b.whs = rdr["prd_whs"].ToString().TrimEnd(' ');
                        b.tag = rdr["prd_tag_no"].ToString().TrimEnd(' ');
                        b.typ = rdr["prd_invt_typ"].ToString();
                        b.sts = rdr["prd_invt_sts"].ToString();
                        b.ref_pfx = rdr["pcr_crtd_ref_pfx"].ToString();
                        b.ref_no= rdr["pcr_crtd_ref_no"].ToString();
                        b.qds = rdr["qds_qds_ctl_no"].ToString().TrimEnd(' ');
                        b.dt = rdr["qds_crtd_dt"].ToString();                

                        lstQDSDataModel.Add(b);
                    }
                }
            }
            catch (OdbcException)
            {
                throw;
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

            return lstQDSDataModel;
        }
    }
}
