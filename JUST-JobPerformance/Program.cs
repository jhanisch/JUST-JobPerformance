using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using JUST.NewJobNotifier.Classes;

namespace JUST.NewJobNotifier
{
    class MainClass
    {
        /* version 1.00a  */
        private const string debug = "debug";
        private const string live = "live";
        private const string monitor = "monitor";

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static string Uid;
        private static string Pwd;
        private static string FromEmailAddress;
        private static string FromEmailPassword;
        private static string FromEmailSMTP;
        private static int? FromEmailPort;
        private static string Mode;
        private static string[] MonitorEmailAddresses;
        private static ArrayList ValidModes = new ArrayList() { debug, live, monitor };

        private static string EmailSubject = "New Jobs created";
        private static string MessageBodyFormat = @"
    <body style = ""margin-left: 20px; margin-right:20px"" >
        <hr/>
        <h2> New Jobs Created</h2>
        <hr/>

        <table style = ""width:80%; text-align: left"" border=""1"" cellpadding=""10"" cellspacing=""0"">
            <tr style = ""background-color: cyan"" >
                <th>Customer</th>
                <th>Customer Name</th>
                <th>Job Number</th>
                <th>Job Description</th>
            </tr>";

        private static string messageBodyTableItem = @"<tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
                <td>{3}</td>
            </tr>";
        private static string messageBodyTail = @"</table></body>";

        static void Main(string[] args)
        {
            try
            {
                log.Info("[Main] Starting up at " + DateTime.Now);

                getConfiguration();

                ProcessNewJobData();

                log.Info("[Main] Completion at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                log.Error("[Main] Error: " + ex.Message);
            }
        }

        private static void getConfiguration()
        {
            Uid = ConfigurationManager.AppSettings["Uid"];
            Pwd = ConfigurationManager.AppSettings["Pwd"];
            FromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"];
            FromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
            FromEmailSMTP = ConfigurationManager.AppSettings["FromEmailSMTP"];
            FromEmailPort = Convert.ToInt16(ConfigurationManager.AppSettings["FromEmailPort"]);

            Mode = ConfigurationManager.AppSettings["Mode"].ToLower();
            var ExecutiveEmailAddressList = ConfigurationManager.AppSettings["ExecutiveEmailAddresses"];
            if (ExecutiveEmailAddressList.Length > 0)
            {
                char[] delimiterChars = { ';', ',' };
                MonitorEmailAddresses = ExecutiveEmailAddressList.Split(delimiterChars);
            }

            #region Validate Configuration Data
            var errorMessage = new StringBuilder();
            if (String.IsNullOrEmpty(Uid))
            {
                errorMessage.Append("User ID (Uid) is Required");
            }

            if (String.IsNullOrEmpty(Pwd))
            {
                errorMessage.Append("Password (Pwd) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailAddress))
            {
                errorMessage.Append("From Email Address (FromEmailAddress) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailPassword))
            {
                errorMessage.Append("From Email Password (FromEmailPassword) is Required");
            }

            if (String.IsNullOrEmpty(FromEmailSMTP))
            {
                errorMessage.Append("From Email SMTP (FromEmailSMTP) address is Required");
            }

            if (!FromEmailPort.HasValue)
            {
                errorMessage.Append("From Email Port (FromEmailPort) is Required");
            }

            if (String.IsNullOrEmpty(Mode))
            {
                errorMessage.Append("Mode is Required");
            }

            if (!ValidModes.Contains(Mode.ToLower()))
            {
                errorMessage.Append(String.Format("{0} is not a valid Mode.  Valid modes are 'debug', 'live' and 'monitor'", Mode));
            }

            if ((Mode == monitor) || (Mode == debug))
            {
                if (MonitorEmailAddresses == null || MonitorEmailAddresses.Length == 0)
                {
                    errorMessage.Append("Executive Email Address is Required in debug & monitor mode");
                }
            }

            if (errorMessage.Length > 0)
            {
                throw new Exception(errorMessage.ToString());
            }
            #endregion
        }

        private static void ProcessNewJobData()
        {
            try
            {
                OdbcConnection cn;
                OdbcCommand cmd;
                var notifiedlist = new ArrayList();
                //var runDate = DateTime.Today.AddDays(1);
                var runDate = new DateTime(2018, 06, 6);
                var checkDate = String.Format("{0:yyyy-MM-dd}", runDate);
                //checkDate = "2018-05-31";
                log.Info("checkDate: " + checkDate);

                //jcjob
                // user_1 = job description
                // user_2 = sales person
                // user_3 = designer
                // user_4 = project manager
                // user_5 = SM
                // user_6 = Fitter
                // user_7 = Plumber
                // user_8 = Tech 1
                // user_9 = Tech 2
                // user_10 = Notified
                //
                //customer
                // user_1 = primary contact
                // user_2 = secondary contact
                var JobsQuery = "select distinct jobnum from prledgerjc where checkdate = {d'" + checkDate + "'} and jobnum is not null and (jobnum like 'C%' or jobnum like 'P%')";

                OdbcConnectionStringBuilder just = new OdbcConnectionStringBuilder();
                just.Driver = "ComputerEase";
                just.Add("Dsn", "Company 0");
                just.Add("Uid", Uid);
                just.Add("Pwd", Pwd);

                cn = new OdbcConnection(just.ConnectionString);
                cmd = new OdbcCommand(JobsQuery, cn);
                cn.Open();
                log.Info("[ProcessNewJobsData] Connection to database opened successfully");

                OdbcDataReader reader = cmd.ExecuteReader();
                try
                {
                    var EmployeeEmailAddresses = GetEmployees(cn);
                    var jobNumColumn = reader.GetOrdinal("jobnum");
                    var executiveNewJobNotifications = new List<JobInformation>();


                    while (reader.Read())
                    {
                        var jobNumber = reader.GetString(jobNumColumn);
                        log.Info("----------------- Found Job Number " + jobNumber + " -------------------");
                        var totalActualHoursForJob = GetTotalActualHoursForJob(cn, jobNumber);
                        var totalEstimatedHoursForJob = GetTotalEstimatedHoursForJob(cn, jobNumber);
                        var x = GetHoursEnteredDuringPayPeriod(cn, jobNumber, runDate);

                        log.Info("    totalActualHoursForJob: " + totalActualHoursForJob.ToString());
                        log.Info("    totalEstimatedHoursForJob: " + totalEstimatedHoursForJob.ToString());

                        if ((Mode == monitor) || (Mode == debug))
                        {
//                            executiveNewJobNotifications.Add(new JobInformation() {JobNumber = jobNumber, JobName = jobName, CustomerNumber = customerNumber, CustomerName = customerName});
                        }
                    }
                    /*
                    foreach (var emp in EmployeeEmailAddresses)
                    {
                        if (emp.NewJobs.Count > 0)
                        {
                            log.Info(" emp: " + emp.Name + ", " + emp.EmailAddress + ", newJobs: " + emp.NewJobs.Count());
                            var message = MessageBodyFormat;
                            foreach (var job in emp.NewJobs)
                            {
                                message += string.Format(messageBodyTableItem, job.CustomerNumber, job.CustomerName, job.JobNumber, job.JobName);
                            }

                            message += messageBodyTail;

                            if ((Mode == live) || (Mode == monitor))
                            {
                                if (sendEmail(emp.EmailAddress, EmailSubject, message))
                                {
                                    foreach (var job in emp.NewJobs)
                                    {
                                        if (!notifiedlist.Contains(job.JobNumber))
                                        {
                                            notifiedlist.Add(job.JobNumber);
                                        }
                                    }
                                }
                                log.Info(message);
                            }
                            else
                            {
                                log.Info(" email would have been sent to " + emp.EmailAddress + ", \r\n" + message);
                            }
                        }
                    }

                    log.Info(" ExecutiveNewJobNotifications: " + executiveNewJobNotifications.Count());
                    if (((Mode == monitor) || (Mode == debug)) && (executiveNewJobNotifications.Count() > 0))
                    {
                        var excutiveMessage = MessageBodyFormat;
                        foreach (var job in executiveNewJobNotifications)
                        {
                            excutiveMessage += string.Format(messageBodyTableItem, job.CustomerNumber, job.CustomerName, job.JobNumber, job.JobName);
                        }

                        excutiveMessage += messageBodyTail;

                        foreach(var executive in MonitorEmailAddresses)
                        {
                            if (sendEmail(executive, EmailSubject, excutiveMessage))
                            {
                                foreach (var job in executiveNewJobNotifications)
                                {
                                    if (!notifiedlist.Contains(job.JobNumber))
                                    {
                                        notifiedlist.Add(job.JobNumber);
                                    }
                                }
                            }
                        }
                    }
                    */
                }
                catch (Exception x)
                {
                    log.Error("[ProcessNewJobsData] Reader Error: " + x.Message);
                }
/*
                foreach (var jobNum in notifiedlist)
                {
                    try
                    {
                        var updateCommand = string.Format("update jcjob set \"user_10\" = 1 where jcjob.jobnum = '{0}'", jobNum);
                        cmd = new OdbcCommand(updateCommand, cn);
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception x)
                    {
                        log.Error(String.Format("[ProcessNewJobsData] Error updating Job Number {0} to be Notified: {1}", jobNum, x.Message));
                    }
                }
*/
                reader.Close();
                cn.Close();
            }
            catch (Exception x)
            {
                log.Error("[ProcessNewJobsData] Exception: " + x.Message);
                return;
            }

            return;
        }

        private static Employee GetEmployeeInformation(List<Employee> EmployeeEmailAddresses, string employee)
        {
            try
            {
                var e = EmployeeEmailAddresses.FirstOrDefault(x => x.EmployeeId.ToLowerInvariant() == employee.ToLowerInvariant());

                if (e != null && e.EmailAddress.Length > 0)
                {
                    return e;
                }
            }
            catch (KeyNotFoundException)
            {
                log.Info("[GetEmployeeInformation] No Employee record found by employeeid for : " + employee);
            }
            catch (Exception x)
            {
                log.Error("[GetEmployeeInformation] by employeeid exception: " + x.Message);
            }

            try
            {
                var e = EmployeeEmailAddresses.FirstOrDefault(x => x.Name.ToLowerInvariant() == employee.ToLowerInvariant());

                if (e != null && e.EmailAddress.Length > 0)
                {
                    return e;
                }
            }
            catch (KeyNotFoundException)
            {
                log.Info("[GetEmployeeInformation] No Employee record found by name for : " + employee);
            }
            catch (Exception x)
            {
                log.Error("[GetEmployeeInformation] by name exception: " + x.Message);
            }

            return new Employee();
        }

        private static string FormatEmailBody(string receivedOnDate, string purchaseOrderNumber, string receivedBy, string bin, string buyerName, string vendor, JobInformation job, string notes)
        {
            var purchaseOrderItemTable = string.Empty;
            foreach (PurchaseOrderItem poItem in job.PurchaseOrderItems)
            {
                purchaseOrderItemTable += string.Format(messageBodyTableItem, poItem.ItemNumber, poItem.Description, poItem.Quantity);
            }

            var emailBody = String.Format(MessageBodyFormat, purchaseOrderNumber, receivedBy, receivedOnDate, bin, job.JobNumber, job.JobName, job.CustomerName, vendor, buyerName, notes) + purchaseOrderItemTable + messageBodyTail;

            return emailBody;
        }

        private static bool sendEmail(string toEmailAddress, string subject, string emailBody)
        {
            return false;

            bool result = true;
            if (toEmailAddress.Length == 0)
            {
                log.Error("  [sendEmail] No toEmailAddress to send message to");
                return false;
            }

            log.Info("  [sendEmail] Sending Email to: " + toEmailAddress);
            log.Info("  [sendEmail] EmailMessage: " + emailBody);

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(FromEmailAddress, "New Job Notification");
                    mail.To.Add(toEmailAddress);
                    mail.Subject = subject;
                    mail.Body = emailBody;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient(FromEmailSMTP, FromEmailPort.Value))
                    {
                        smtp.Credentials = new NetworkCredential(FromEmailAddress, FromEmailPassword);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                        log.Info("  [sendEmail] Email Sent to " + toEmailAddress);
                    }
                }
            }
            catch (Exception x)
            {
                result = false;
                log.Error(String.Format("  [sendEmail] Error Sending email to {0}, message: {1}", x.Message, emailBody));
            }

            return result;
        }

        private static long GetTotalActualHoursForJob(OdbcConnection cn, string jobNumber)
        {
            var jobHours = 0;
            var jobQuery = "Select sum(jcdetail.hours) from jcdetail where jobnum = '{0}'";
            var buyerCmd = new OdbcCommand(string.Format(jobQuery, jobNumber), cn);

            OdbcDataReader jobReader = buyerCmd.ExecuteReader();

            if (jobReader.Read())
            {
                jobHours = jobReader.GetInt32(0);
            }

            jobReader.Close();

            return jobHours;
        }

        private static long GetTotalEstimatedHoursForJob(OdbcConnection cn, string jobNumber)
        {
            var estimatedJobHours = 0;
            var jobQuery = "Select sum(jccat.budget1_hours) from jccat where jobnum = '{0}'";
            var buyerCmd = new OdbcCommand(string.Format(jobQuery, jobNumber), cn);

            OdbcDataReader jobReader = buyerCmd.ExecuteReader();

            if (jobReader.Read())
            {
                estimatedJobHours = jobReader.GetInt32(0);
            }

            jobReader.Close();

            return estimatedJobHours;
        }

        private static long GetHoursEnteredDuringPayPeriod(OdbcConnection cn, string jobNumber, DateTime checkDate)
        {
            long offset = 0;

            if (checkDate.DayOfWeek == DayOfWeek.Monday) {
                offset = -9;
            }
            else if (checkDate.DayOfWeek == DayOfWeek.Tuesday)
            {
                offset = -10;
            }
            else if (checkDate.DayOfWeek == DayOfWeek.Wednesday)
            {
                offset = -11;
            }
            else if (checkDate.DayOfWeek == DayOfWeek.Thursday)
            {
                offset = -12;
            }

            var periodStartDate = checkDate.AddDays(offset);
            var periodEndDate = checkDate.AddDays(offset + 8);

            var formattedPeriodStartDate = String.Format("{0:yyyy-MM-dd}", periodStartDate);
            var formattedPeriodEndDate = String.Format("{0:yyyy-MM-dd}", periodEndDate);

            var jobQuery = "Select who, date, hours, type from jcdetail where jobnum = '"+ jobNumber + "' and date > {d'" + formattedPeriodStartDate + "'} and date < {d'" + formattedPeriodEndDate + "'} order by who, date asc";
            var buyerCmd = new OdbcCommand(jobQuery, cn);

            OdbcDataReader hoursReader = buyerCmd.ExecuteReader();

            try
            {
                var whoColumn = hoursReader.GetOrdinal("who");
                var dateColumn = hoursReader.GetOrdinal("date");
                var hoursColumn = hoursReader.GetOrdinal("hours");
                var typeColumn = hoursReader.GetOrdinal("type");

                while (hoursReader.Read())
                {
                    var hours = hoursReader.GetDecimal(hoursColumn);
                    if (hours > 0) {
                        var who = hoursReader.GetString(whoColumn);
                        var date = hoursReader.GetDateTime(dateColumn);
                        var type = hoursReader.GetInt32(typeColumn);
                        log.Info("  " + who + ": " + hours.ToString() + " type " + type.ToString() + " hours on " + date.ToShortDateString());
                   }
                }
            }
            catch (Exception ex)
            {
                log.Error("GetHoursEnteredDuringPayPeriod error: " + ex.Message);
            }


            hoursReader.Close();

            return 0;
        }



        private static List<Employee> GetEmployees(OdbcConnection cn)
        {
            var employees = new List<Employee>();

            var buyerQuery = "Select user_1, user_2, name from premployee where user_1 is not null";
            var buyerCmd = new OdbcCommand(buyerQuery, cn);

            OdbcDataReader buyerReader = buyerCmd.ExecuteReader();

            try
            {
                while (buyerReader.Read())
                {
                    var buyer = buyerReader.GetString(0);
                    var email = buyerReader.GetString(1);
                    var name = buyerReader.GetString(2);

                    if (buyer.Trim().Length > 0)
                    {
                        employees.Add(new Employee(buyer, name, email));
                    }
                }

            }
            catch (Exception ex)
            {
                log.Error("GetEmployees Exception: " + ex.Message);
            }

            buyerReader.Close();

            return employees;
        }
    }
}
/*
log.Info("column names");
for (int col = 0; col < reader.FieldCount; col++)
{
    log.Info(reader.GetName(col));
}
*/
