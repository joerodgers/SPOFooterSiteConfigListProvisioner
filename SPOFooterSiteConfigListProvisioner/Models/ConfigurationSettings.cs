namespace SPOFooterSiteConfigListProvisioner
{
    public class ConfigurationSettings
    {
        public const string SectionName = "ConfigurationSettings";

        public string CertificatePassword
        {
            get;
            set;
        }

        public string CertificatePath
        {
            get;
            set;
        }

        public string ClientId
        {
            get;
            set;
        }

        public bool Force
        {
            get;
            set;
        }

        public int MaximumThreads
        {
            get;
            set;
        } = 10;

        public string SecurityGroupName
        {
            get;
            set;
        }

        public string Tenant
        {
            get;
            set;
        }

        public string TenantUrl => $"https://{Tenant}-admin.sharepoint.com";

        public string AzureTenant => $"{Tenant}.onmicrosoft.com";
    }
}
