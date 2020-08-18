using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Text;

namespace SPOFooterSiteConfigListProvisioner.Interfaces
{
    public interface ISiteCollectionExecuter
    {
        void Execute(string url);
    }
}
