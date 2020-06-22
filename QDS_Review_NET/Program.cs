using System;
using System.Net.Mail;
using System.Configuration;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using STXtoSQL.Log;
using QDS_Review_NET.Models;
using QDS_Review_NET.DataAccess;

namespace QDS_Review_NET
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger.LogWrite("MSG", "Start: " + DateTime.Now.ToString());

            // Args initialization  
            bool emailIt = false;

            /*
             * Copy empty File from templates to destination folder
             * Create a full path to pass to Excel methods
             */
            string fileName = ConfigurationManager.AppSettings.Get("FileName");
            string templatePath = ConfigurationManager.AppSettings.Get("TemplatePath");
            string destPath = ConfigurationManager.AppSettings.Get("DestPath");

            File.Copy(Path.Combine(templatePath, fileName), Path.Combine(destPath, fileName), true);

            // Pass to Excel methods to be used in OleDb connection string
            string fullPath = Path.Combine(destPath, fileName);

            #region Args
            /*
             * arg options:
             * 1. Email = true or false
             */
            try
            {
                // More than 1 arg[] is invalid, else get Email value from arg[0]
                if ((args.Length != 1))
                {
                    Logger.LogWrite("MSG", "Invalid number of args[]");
                    Logger.LogWrite("MSG", "Return on args[]");
                    return;
                }
                else
                {
                    emailIt = Convert.ToBoolean(args[0]);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Args");
                return;
            }
            #endregion

            #region ODBCData
            ODBCData objQDS = new ODBCData();

            List<QDSDataModel> lstLogData = new List<QDSDataModel>();

            try
            {
                lstLogData = objQDS.Get_QDSNotAppv();
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on LogData");
                return;
            }
            #endregion

            #region Excel          
            /*
             * Export each List<object> to correct XLS tab
             */
            ExcelExport objXLS = new ExcelExport();

            // Log - Tab
            try
            {
                objXLS.WriteQDSNotAppv(lstLogData, fullPath);
            }
            catch (Exception ex)
            {
                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return on Bookings XLS");
                return;
            }
            #endregion

            #region Email
            /*
             * Emailing the file is optoinal
             * emailIt = true or false           
             */
            if (emailIt)
            {
                List<EmployeesReportsModel> lstEmpReports = new List<EmployeesReportsModel>();

                SQLData objSQL_Rpts = new SQLData();

                try
                {
                    lstEmpReports = objSQL_Rpts.Get_Emp_Reports("SW", "QDSDaily");
                }
                catch (Exception ex)
                {
                    Logger.LogWrite("EXC", ex);
                    Logger.LogWrite("MSG", "Return on Emp Rpts");
                    return;
                }

                try
                {
                    string emailTo = ConfigurationManager.AppSettings.Get("EmailTo");
                    string pwd365 = Environment.GetEnvironmentVariable("sysOffice365PWD");

                    MailMessage mail = new MailMessage();

                    SmtpClient SmtpServer = new SmtpClient("smtp.office365.com");

                    mail.From = new MailAddress("noreply@calstripsteel.com");
                    mail.Subject = "QDS - Daily";
                    mail.Body = "Report attached";

                    //Build To: line from emails in list of EmployeesReportsModel
                    foreach (EmployeesReportsModel e in lstEmpReports)
                    {
                        Logger.LogWrite("MSG", "Email: " + e.email.ToString());
                        mail.To.Add(e.email.ToString());
                    }

                    // Add attachment
                    Attachment attach;
                    attach = new Attachment(fullPath);
                    mail.Attachments.Add(attach);

                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("noreply@calstripsteel.com", pwd365);
                    SmtpServer.EnableSsl = true;

                    SmtpServer.Send(mail);

                    Logger.LogWrite("MSG", "Email: " + emailTo);
                }
                catch (Exception ex)
                {
                    Logger.LogWrite("EXC", ex);
                    Logger.LogWrite("MSG", "Return on Email");
                    return;
                }
            }
            else
                Logger.LogWrite("MSG", "No email");
            #endregion

            // Made it to the end
            Logger.LogWrite("MSG", "End: " + DateTime.Now.ToString());

            // Testing
            //Console.WriteLine("Press key to exit");
            //Console.ReadKey();
        }
    }
}
