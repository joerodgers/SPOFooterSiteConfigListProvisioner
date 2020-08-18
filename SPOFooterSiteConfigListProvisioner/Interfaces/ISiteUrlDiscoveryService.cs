using System;
using System.Collections.Generic;
using System.Text;

namespace SPOFooterSiteConfigListProvisioner.Interfaces
{
    public interface ISiteUrlDiscoveryService
    {
        List<string> GetSiteCollectionUrls();
    }
}
