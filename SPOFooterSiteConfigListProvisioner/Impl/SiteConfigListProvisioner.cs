using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using SPOFooterSiteConfigListProvisioner.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SPOFooterSiteConfigListProvisioner
{
    public class SiteConfigListProvisioner : ISiteCollectionExecuter
    {
        private static readonly string _listTitle       = "SiteConfig";
        private static readonly string _listRelativeUrl = $"Lists/{_listTitle}";
        private static readonly string _sponsorTitle    = "SITE_SPONSOR";
        private static readonly string _ownerTitle      = "SITE_PRIMARY_ADMIN";
        private static readonly string _fieldValueName  = "Value";

        private readonly ILogger<SiteConfigListProvisioner> _logger;
        private readonly IOptions<ConfigurationSettings> _settings;

        public SiteConfigListProvisioner(ILogger<SiteConfigListProvisioner> logger, IOptions<ConfigurationSettings> settings)
        {
            _logger   = logger   ?? throw new ArgumentNullException(nameof(logger));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public void Execute(string url)
        {
            _logger.LogInformation($"Processing {url}");
            ProvisionSiteConfigList(url);
        }

        private void ProvisionSiteConfigList(string url)
        {
            var ctx = CreateClientContext(url);

            if (ctx != null)
            {
                try
                {
                    ctx.Load(ctx.Site, prop => prop.Owner);
                    ctx.Load(ctx.Web, prop => prop.RoleDefinitions,
                                      prop => prop.Url,
                                      prop => prop.AssociatedMemberGroup);

                    var list = ctx.Web.GetListByUrl(_listRelativeUrl, i => i.Fields,
                                                                      i => i.ParentWeb,
                                                                      i => i.DefaultViewUrl,
                                                                      i => i.HasUniqueRoleAssignments);

                    if (list == null)
                    {
                        list = ProvisionList(ctx.Web);

                        if (list != null)
                        {
                            SetListPermissions(list);
                            ProvisionFields(list);
                            ProvisionDefaultItem(list);
                        }
                    }
                    else if (_settings.Value.Force)
                    {
                        SetListPermissions(list);
                        ProvisionFields(list);
                        ProvisionDefaultItem(list);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Failed to create list for {url}");
                }
                finally
                {
                    ctx.Dispose();
                }

                return;
            }

            _logger.LogError($"Client context was null for {url}");
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

        private void ProvisionDefaultItem(List list)
        {
            // load all list items, which should only be a handful at most
            var listItems = list.GetItems(CamlQuery.CreateAllItemsQuery());
            var cc = list.Context as ClientContext;

            list.Context.Load(listItems, props => props.Include(itemprops => itemprops.Client_Title, itemprops => itemprops.FieldValuesForEdit));
            list.Context.Load(cc.Site, props => props.Owner);
            list.Context.ExecuteQueryRetry();

            // check if there are any items for the sponsor, create if missing
            if (!listItems.Where(w => w.Client_Title == _sponsorTitle).Any())
            {
                _logger.LogDebug($"Creating {_sponsorTitle} item in {list.DefaultViewUrl}");

                var item = list.AddItem(new ListItemCreationInformation());
                item["Title"] = _sponsorTitle;
                item["Value"] = string.Empty;
                item.Update();
            }

            // check if there are any items for the primary admin, create if missing
            var primaryOwnerItem = listItems.Where(w => w.Client_Title == _ownerTitle).FirstOrDefault();

            if (primaryOwnerItem == null)
            {
                // create new item
                _logger.LogDebug($"Creating {_ownerTitle} item in {list.DefaultViewUrl}");

                primaryOwnerItem = list.AddItem(new ListItemCreationInformation());
                primaryOwnerItem["Title"] = _ownerTitle;
                primaryOwnerItem["Value"] = cc.Site.Owner.LoginName;
                primaryOwnerItem.Update();
            }
            else if (primaryOwnerItem.FieldValuesForEdit["Value"] != cc.Site.Owner.LoginName)
            {
                // update existing item
                _logger.LogDebug($"Setting {_ownerTitle} value to {((ClientContext)list.Context).Site.Owner.LoginName}");

                primaryOwnerItem["Value"] = cc.Site.Owner.LoginName;
                primaryOwnerItem.Update();
            }

            // commit changes to SPO
            if (list.Context.HasPendingRequest)
            {
                listItems = list.GetItems(CamlQuery.CreateAllItemsQuery());
                list.Context.Load(listItems, x => x.Include(y => y.Client_Title, y => y.FieldValuesForEdit));
                list.Context.ExecuteQueryRetry();
            }

            // make sure AAD group perm is set on sponsor list item
            var sponsorItem = listItems.Where(w => w.Client_Title == _sponsorTitle).FirstOrDefault();
            if (sponsorItem != null && !string.IsNullOrWhiteSpace(_settings.Value.SecurityGroupName))
            {
                var principal = list.ParentWeb.EnsureUser(_settings.Value.SecurityGroupName);
                list.Context.ExecuteQueryRetry();

                if (principal != null)
                {
                    _logger.LogDebug($"Adding {_settings.Value.SecurityGroupName} group to {_sponsorTitle} in {list.DefaultViewUrl}");

                    sponsorItem.BreakRoleInheritance(true, true);
                    sponsorItem.Update();

                    var roleAssignments = sponsorItem.RoleAssignments.Add(
                        principal,
                        new RoleDefinitionBindingCollection(list.Context)
                        {
                            list.ParentWeb.RoleDefinitions.GetByName("Contribute")
                        });

                    list.Context.Load(roleAssignments);
                    list.Context.ExecuteQueryRetry();
                }
                else
                {
                    _logger.LogWarning($"Group not added to site: {_settings.Value.SecurityGroupName}");
                }
            }
        }

        private void ProvisionFields(List list)
        {
            list.Context.Load(list.Fields, f => f.Include(i => i.InternalName));
            list.Context.ExecuteQueryRetry();

            if (list.Fields.Where(w => w.InternalName == _fieldValueName).Any())
            {
                _logger.LogDebug($"{_fieldValueName} field exists on {list.DefaultViewUrl}");
            }
            else
            {
                _logger.LogDebug($"Creating {_fieldValueName} field on {list.DefaultViewUrl}");

                var fieldCreationInformation = new FieldCreationInformation(FieldType.Note)
                {
                    Id               = Guid.NewGuid(),
                    InternalName     = _fieldValueName,
                    DisplayName      = _fieldValueName,
                    AddToDefaultView = true
                };

                list.CreateField(fieldCreationInformation, true);
            }
        }

        private List ProvisionList(Web web)
        {
            _logger.LogDebug($"Provisioning List on {web.Url}");

            // create the config list
            var list = web.Lists.Add(new ListCreationInformation()
            {
                Title             = _listTitle,
                TemplateType      = (int)ListTemplateType.GenericList,
                QuickLaunchOption = QuickLaunchOptions.Off
            });

            list.Hidden = true;
            list.EnableVersioning = true;
            list.NoCrawl = true;
            list.Update();

            web.Context.Load(list,
                i => i.Fields,
                i => i.ParentWeb,
                i => i.DefaultViewUrl,
                i => i.HasUniqueRoleAssignments);

            list.Context.ExecuteQueryRetry();

            return list;
        }

        private void SetListPermissions(List list)
        {
            if (!list.HasUniqueRoleAssignments || !_settings.Value.Force)
            {
                var membersGroup = list.ParentWeb.AssociatedMemberGroup;

                if (membersGroup == null)
                {
                    _logger.LogWarning($"Members Group not found on {list.DefaultViewUrl}");
                    return;
                }

                _logger.LogDebug($"Granting Members Group Read to {list.DefaultViewUrl}");

                // ensure list perms are broken
                list.BreakRoleInheritance(true, true);

                // add "read" perms to the members group
                {
                    var roleDefinition = new RoleDefinitionBindingCollection(list.Context)
                    {
                        list.ParentWeb.RoleDefinitions.GetByName("Read")
                    };

                    list.Context.Load(list.RoleAssignments.Add(membersGroup, roleDefinition));
                    list.Context.ExecuteQueryRetry();
                }

                _logger.LogDebug($"Removing Members Edit rights from {list.DefaultViewUrl}");

                // remove "edit" perms from the members group
                {
                    var roleAssignment = list.RoleAssignments.GetByPrincipal(membersGroup);
                    var roleDefinitionBindings = roleAssignment.RoleDefinitionBindings;

                    list.Context.Load(roleDefinitionBindings);
                    list.Context.ExecuteQueryRetry();

                    foreach (var roleDefinition in roleDefinitionBindings.Where(roleDefinition => roleDefinition.Name == "Edit"))
                    {
                        roleDefinitionBindings.Remove(roleDefinition);
                        roleAssignment.Update();
                        list.Context.ExecuteQueryRetry();
                        break;
                    }
                }
            }
            else
            {
                _logger.LogDebug($"HasUniqueRoleAssignments == true for {list.DefaultViewUrl}");
            }
        }
    }
}
