using System;
namespace JUST.PONotifier.Classes
{
    public class JobInformation
    {
        public JobInformation() { }
        public JobInformation(string projectManagerName, string jobNumber, string jobName)
        {
            ProjectManagerName = projectManagerName;
            JobNumber = jobNumber;
            JobName = jobName;
        }

        public string ProjectManagerName { get; set; }
        public string JobNumber { get; set; }
        public string JobName { get; set; }
    }
}
