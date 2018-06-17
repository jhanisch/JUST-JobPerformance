using System;
using System.Collections;
using System.Collections.Generic;

namespace JUST.JobPerformanceNotifier.Classes
{
    public class Employee
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Employee()
        {
            EmployeeId = string.Empty;
            Name = string.Empty;
            EmailAddress = string.Empty;
            JobPerformanceMessage = new List<String>();
        }

        public Employee(string employeeId, string name, string emailAddress)
        {
            EmployeeId = employeeId;
            Name = name;
            EmailAddress = emailAddress;
            JobPerformanceMessage = new List<String>();
        }

        public void AddMessageToNotify(string message)
        {
            log.Info("Adding " + message + ": to " + Name);
            JobPerformanceMessage.Add(message);
        }

        public string EmployeeId { get; set; }
        public string Name { get; set; }
        public string EmailAddress { get; set; } 
        public List<String> JobPerformanceMessage { get; }
    }

    public class JobInformation
    {
        public JobInformation()
        {
            ProjectManagerName = string.Empty;
            JobNumber = string.Empty;
            JobName = string.Empty;
            CustomerNumber = string.Empty;
            CustomerName = string.Empty;
        }

        public JobInformation(string projectManagerName, string jobNumber, string jobName, string customerNumber)
        {
            ProjectManagerName = projectManagerName;
            JobNumber = jobNumber;
            JobName = jobName;
            CustomerNumber = customerNumber;
        }

        public string ProjectManagerName { get; set; }
        public string JobNumber { get; set; }
        public string JobName { get; set; }
        public string CustomerNumber { get; set; }
        public string CustomerName { get; set; }
    }

    public class JobDetail
    {
        public JobDetail()
        {
            Who = String.Empty;
            Type = null;
            Hours = null;
        }

        public JobDetail(string who, long type, decimal hours) 
        {
            Who = who;
            Type = type;
            Hours = hours;
        }

        public string Who { get; set; }
        public long? Type { get; set; }
        public decimal? Hours { get; set; }
    }
}
