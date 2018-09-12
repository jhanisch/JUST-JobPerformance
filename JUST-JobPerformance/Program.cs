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
using JUST.JobPerformanceNotifier.Classes;

namespace JUST.JobPerformanceNotifier
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

        private static string EmailSubject = "Weekly Jobs Performance Report: {0}";
        private static string MessageBodyHeader = @"
<body style = ""margin-left: 20px; margin-right:20px"" >";
        private static string MessageBodyFormat = @"<hr/>
        <h2> Performance Summary for Job {0} - {1}</h2>
        <p>
        <strong>Primary Contact:</strong> {2}<br/>
        <strong>Customer:</strong> {3} - {4}<br/>
        <strong>Check Date:</strong> {5}<br/>
        </p>

        <table style = ""width:80%; text-align: left"" border=""1"" cellpadding=""10"" cellspacing=""0"">
            <tr style = ""background-color: cyan"" >
                <th>Technician</th>
                <th>This Weeks Hours</th>
                <th>Hour Type</th>
            </tr>";

        private static string MessageBodyTableItem = @"<tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
            </tr>";
        private static string MessageBodyTail = @"</table>
        <br>
        <p>
            <strong>Actual Total Job Hours:</strong> {0}<br/>
            <strong>Estimated Total Hours:</strong> {1}<br/>
            <strong>Remaining:</strong> {2}<br/>
        </p>
        <br>";
        private static string EndMessageBody = "</body>";

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
                var runDate = DateTime.Today; 
                runDate = runDate.AddDays(runDate.DayOfWeek == DayOfWeek.Monday ? 2 : 1);
                //runDate = new DateTime(2018, 06, 20);
                var checkDate = String.Format("{0:yyyy-MM-dd}", runDate);
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
//                var JobsQuery = "select distinct prledgerjc.jobnum, customer.user_1 from prledgerjc inner join jcjob on prledgerjc.jobnum = jcjob.jobnum  inner join customer on jcjob.cusnum = customer.cusnum where checkdate = {d'" + checkDate + "'} and prledgerjc.jobnum is not null and (prledgerjc.jobnum like 'C%' or prledgerjc.jobnum like 'P%') and prledgerjc.jobnum = 'CAC18081'";
                var JobsQuery = "select distinct prledgerjc.jobnum, jcjob.user_2 as primaryContact, jcjob.name as jobName, customer.cusnum as customerNumber, customer.name as customerName from prledgerjc inner join jcjob on prledgerjc.jobnum = jcjob.jobnum  inner join customer on jcjob.cusnum = customer.cusnum where checkdate = {d'" + checkDate + "'} and prledgerjc.jobnum is not null and (prledgerjc.jobnum like 'C%' or prledgerjc.jobnum like 'P%')";

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
                    var jobNameColumn = reader.GetOrdinal("jobName");
                    var customerNumberColumn = reader.GetOrdinal("customerNumber");
                    var customerNameColumn = reader.GetOrdinal("customerName");
                    var primaryContactColumn = reader.GetOrdinal("primaryContact");
                    var executiveNewJobNotifications = new List<String>();

                    while (reader.Read())
                    {
                        var jobNumber = reader.GetString(jobNumColumn);
                        var jobName = reader.GetString(jobNameColumn);
                        var primaryContact = reader.GetString(primaryContactColumn);
                        var customerNumber = reader.GetString(customerNumberColumn);
                        var customerName = reader.GetString(customerNameColumn);
                        log.Info("\r\n----------------- Found Job Number " + jobNumber + " for contact " + primaryContact + ": " + GetEmployeeInformation(EmployeeEmailAddresses, primaryContact).EmailAddress + " -------------------");
                        var totalActualHoursForJob = GetTotalActualHoursForJob(cn, jobNumber);
                        var totalEstimatedHoursForJob = GetTotalEstimatedHoursForJob(cn, jobNumber);
                        var x = GetHoursEnteredDuringPayPeriod(cn, jobNumber, runDate, totalActualHoursForJob, totalEstimatedHoursForJob);
                        log.Info("    totalActualHoursForJob: " + totalActualHoursForJob.ToString());
                        log.Info("    totalEstimatedHoursForJob: " + totalEstimatedHoursForJob.ToString());

                        var emp = GetEmployeeInformation(EmployeeEmailAddresses, primaryContact);
                        var message = string.Format(MessageBodyFormat, jobNumber, jobName, emp.Name, customerNumber, customerName, checkDate) + x;
                        emp.AddMessageToNotify(message);

                        if ((Mode == monitor) || (Mode == debug))
                        {
                            executiveNewJobNotifications.Add(message);
                        }
                    }

                    if ((Mode == live) || (Mode == monitor))
                    {
                        foreach (var emp in EmployeeEmailAddresses)
                        {
                            if (emp.JobPerformanceMessage.Count > 0)
                            {
                                var r = MessageBodyHeader;
                                foreach (var jobMessage in emp.JobPerformanceMessage)
                                {
                                    r += jobMessage;
                                }

                                r += EndMessageBody;

                                var emailSubject = string.Format(EmailSubject, checkDate);
                                sendEmail(emp.EmailAddress, emailSubject, r);
                            }
                        }
                    }

                    if (((Mode == monitor) || (Mode == debug)) && (executiveNewJobNotifications.Count() > 0))
                    {
                        var executiveMessage = MessageBodyHeader;
                        foreach (var jobMessage in executiveNewJobNotifications)
                        {
                            executiveMessage += jobMessage;
                        }

                        executiveMessage += EndMessageBody;

                        var emailSubject = string.Format(EmailSubject, checkDate);
                        foreach (var executive in MonitorEmailAddresses)
                        {
                            sendEmail(executive, emailSubject + " - Executive", executiveMessage);
                        }
                    }
                }
                catch (Exception x)
                {
                    log.Error("[ProcessNewJobsData] Reader Error: " + x.Message);
                }

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

        private static bool sendEmail(string toEmailAddress, string subject, string emailBody)
        {
            bool result = true;

            if (toEmailAddress.Length == 0)
            {
                log.Error("  [sendEmail] No toEmailAddress to send message to");
                return false;
            }

            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(FromEmailAddress, subject);
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

        private static decimal GetTotalActualHoursForJob(OdbcConnection cn, string jobNumber)
        {
            var jobHours = 0m;
            var jobQuery = "Select sum(jcdetail.hours) from jcdetail where jobnum = '{0}'";
            var buyerCmd = new OdbcCommand(string.Format(jobQuery, jobNumber), cn);

            OdbcDataReader jobReader = buyerCmd.ExecuteReader();

            if (jobReader.Read())
            {
                jobHours = jobReader.GetDecimal(0);
            }

            jobReader.Close();

            return jobHours;
        }

        private static decimal GetTotalEstimatedHoursForJob(OdbcConnection cn, string jobNumber)
        {
            var estimatedJobHours = 0m;
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

        private static string GetHoursEnteredDuringPayPeriod(OdbcConnection cn, string jobNumber, DateTime checkDate, decimal totalActualHours, decimal totalEstimatedHours)
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
            var hoursForJobThisPayPeriod = new List<JobDetail>();

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

                        hoursForJobThisPayPeriod.Add(new JobDetail(who, type, hours));
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("GetHoursEnteredDuringPayPeriod error: " + ex.Message);
            }

            var employeeSummary =
                from h in hoursForJobThisPayPeriod
                group h by new { h.Who, h.Type } into g
                select new { g.Key.Who, g.Key.Type, TotalHours = g.Sum(h => h.Hours) };

            var message = string.Empty;

            foreach (var x in employeeSummary)
            {
                message += String.Format(MessageBodyTableItem, x.Who, x.TotalHours, FormatHoursType(x.Type));
            }

            var totalJobHours = employeeSummary.Sum(h => h.TotalHours);

            if (totalJobHours > 0)
            {
                message += String.Format(MessageBodyTableItem, "<strong>Weekly Total</strong>", totalJobHours.ToString(), string.Empty);
            }

            var remainingHoursForJob = totalEstimatedHours - totalActualHours;
            var formattedRemainingHours = remainingHoursForJob.ToString();
            if (remainingHoursForJob < 0)
            {
                formattedRemainingHours = string.Format(@"<strong><font color=""red"">{0}</font></strong>", remainingHoursForJob);
            }
            message += string.Format(MessageBodyTail, totalActualHours.ToString(), totalEstimatedHours.ToString(), formattedRemainingHours );
            hoursReader.Close();

            return message;
        }

        private static string FormatHoursType(long? HoursType)
        {
            var result = string.Empty;

            if (HoursType.HasValue)
            {
                switch (HoursType)
                {
                    case 1: result = "REG"; break;
                    case 2: result = "OT"; break;
                    default: result = "DBL"; break;
                }
            }

            return result;
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
