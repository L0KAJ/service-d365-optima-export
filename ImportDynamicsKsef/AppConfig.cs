using System;

namespace ImportDynamicsKsef
{
    public class AppConfig
    {
        public AzureAdConfig AzureAd { get; set; }
        public EndpointsConfig Endpoints { get; set; }
        public ConnectionStringsConfig ConnectionStrings { get; set; }
        public SchedulerConfig Scheduler { get; set; }
        public StateConfig State { get; set; }
        public PathsConfig Paths { get; set; }
    }

    public class AzureAdConfig
    {
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
    }

    public class EndpointsConfig
    {
        public string D365Url { get; set; }
        public string JourEntity { get; set; }
        public string TransEntity { get; set; }
    }

    public class ConnectionStringsConfig
    {
        public string Optima { get; set; }
        public string OptimaConfig { get; set; }
    }

    public class SchedulerConfig
    {
        public int TargetHour { get; set; }
        public int TargetMinute { get; set; }
    }

    public class StateConfig
    {
        public DateTime LastRunDate { get; set; }
    }

    public class PathsConfig
    {
        public string ExportFolder { get; set; }
        public string ErrorLog { get; set; }
    }
}