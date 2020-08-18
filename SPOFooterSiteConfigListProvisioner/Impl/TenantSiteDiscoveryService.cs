using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using SPOFooterSiteConfigListProvisioner.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SPOFooterSiteConfigListProvisioner
{
    public class TenantSiteUrlDiscoveryService : ISiteUrlDiscoveryService
    {
        private readonly ILogger<TenantSiteUrlDiscoveryService> _logger;
        private readonly IOptions<ConfigurationSettings> _settings;

        public TenantSiteUrlDiscoveryService(ILogger<TenantSiteUrlDiscoveryService> logger, IOptions<ConfigurationSettings> settings)
        {
            _logger   = logger   ?? throw new ArgumentNullException(nameof(logger));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public List<string> GetSiteCollectionUrls()
        {
            return GetSiteCollectionUrlsInternal();
        }

        private List<string> GetSiteCollectionUrlsInternal()
        {
            var ctx = CreateClientContext(_settings.Value.TenantUrl);
            return new Tenant(ctx).GetSiteCollections(includeDetail: false, includeOD4BSites: false).Select(s => s.Url).ToList();
        }

        private ClientContext CreateClientContext(string url)
        {
            return new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(
                url,
                _settings.Value.ClientId,
                _settings.Value.AzureTenant,
                _settings.Value.CertificatePath,
                _settings.Value.CertificatePassword);
        }
    }
}
