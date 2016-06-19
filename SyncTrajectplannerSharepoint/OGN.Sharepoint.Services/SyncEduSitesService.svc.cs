using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Security;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using System.Diagnostics;
using System.Configuration;
using System.Web.Configuration;

namespace OGN.Sharepoint.Services
{
    public class SyncEduSitesService : ISyncEduSitesService
    {
        //web.config settings (see constructor)
        private string _mod_url;
        private string _edu_url;

        private string _modHome_url;
        private string _eduHome_url;


        private string _edu_siteslist;
        private string _edu_siteslist_column;
        private string _edu_siteslist_school_column;

        private string _mod_siteslist;
        private string _mod_siteslist_column;
        private string _mod_siteslist_school_column;

        private NetworkCredential _creds;
        private Guid _loi_id;
        private Guid _cat_id;
        private Guid _mod_id;
        private Guid _edu_id;
        private string _modtemplate;
        private string _modsubtemplate;
        private string _modsub_title;
        private string _modsub_id;
        private string _moddoclib_berichten;
        private string _moddoclib_examendossier;
        private int _lcid;
        private string _edutemplate;
        private string _edudoclib_berichten;
        private string _edudoclib_examendossier;
        private string _link2edu_list;
        private string _link2edu_list_column;
        private string _link2edu_list_value;
        private string _link2mod_list;
        private string _link2mod_list_column;
        private string _link2mod_list_value;

        private int _modSiteCollectionCount;
        private int _eduSiteCollectionCount;

        private int _modMaxSiteCount;
        private int _eduMaxSiteCount;



        System.Net.Mail.SmtpClient _mailer;
        private string _mailfrom;
        private string _mail2admin;
        private string _mail2business;
        private Configuration configfile;

        //PowerShell: New-EventLog -LogName Application -Source OGN_Sharepoint_Services_SyncEduSitesService
        private string _eventlogsource = "OGN_Sharepoint_Services_SyncEduSitesService";

        public SyncEduSitesService() : this(true) { }
        public SyncEduSitesService(bool isWeb)
        {
            //get web. or app.config settings
            configfile = isWeb
                ? WebConfigurationManager.OpenWebConfiguration("~/")
                : ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            //the credentials of the application pool are used.
            _creds = CredentialCache.DefaultNetworkCredentials; // for production


            _mod_url = ConfigurationManager.AppSettings["sp.sitecollection:mod:url"];
            _edu_url = ConfigurationManager.AppSettings["sp.sitecollection:edu:url"];

            _modHome_url = ConfigurationManager.AppSettings["sp.sitecollection:mod:HomeUrl"];
            _eduHome_url = ConfigurationManager.AppSettings["sp.sitecollection:edu:HomeUrl"];


            _modSiteCollectionCount = int.Parse(ConfigurationManager.AppSettings["sp.sitecollection:mod:count"]);
            _eduSiteCollectionCount = int.Parse(ConfigurationManager.AppSettings["sp.sitecollection:edu:count"]);

            _modMaxSiteCount = int.Parse(ConfigurationManager.AppSettings["sp.sitecollection:mod:maxSiteCount"]);
            _eduMaxSiteCount = int.Parse(ConfigurationManager.AppSettings["sp.sitecollection:edu:maxSiteCount"]);

            _edu_siteslist = ConfigurationManager.AppSettings["sp.sitecollection:edu:list2sites"];
            _edu_siteslist_column = ConfigurationManager.AppSettings["sp.sitecollection:edu:list2sites:column"];
            _edu_siteslist_school_column = ConfigurationManager.AppSettings["sp.sitecollection:edu:list2sites:schoolcolumn"];

            _mod_siteslist = ConfigurationManager.AppSettings["sp.sitecollection:mod:list2sites"];
            _mod_siteslist_column = ConfigurationManager.AppSettings["sp.sitecollection:mod:list2sites:column"];
            _mod_siteslist_school_column = ConfigurationManager.AppSettings["sp.sitecollection:mod:list2sites:schoolcolumn"];

            _lcid = Int32.Parse(ConfigurationManager.AppSettings["sp.site:lcid"]);
            _loi_id = new Guid(ConfigurationManager.AppSettings["sp.termstore:id"]);
            _cat_id = new Guid(ConfigurationManager.AppSettings["sp.termstore.termset:id"]);
            _mod_id = new Guid(ConfigurationManager.AppSettings["sp.termstore.termset.modset:id"]);
            _edu_id = new Guid(ConfigurationManager.AppSettings["sp.termstore.termset.eduset:id"]);
            _modtemplate = ConfigurationManager.AppSettings["sp.modsite:template"];
            _link2edu_list = ConfigurationManager.AppSettings["sp.modsite:list2edu"];
            _link2edu_list_column = ConfigurationManager.AppSettings["sp.modsite:list2edu:column"];
            _link2edu_list_value = ConfigurationManager.AppSettings["sp.modsite:list2edu:value"];
            _modsubtemplate = ConfigurationManager.AppSettings["sp.modsite.subsite:template"];
            _modsub_title = ConfigurationManager.AppSettings["sp.modsite.subsite:title"];
            _modsub_id = ConfigurationManager.AppSettings["sp.modsite.subsite:id"];
            _moddoclib_berichten = ConfigurationManager.AppSettings["sp.modsite:doclib:berichten"];
            _moddoclib_examendossier = ConfigurationManager.AppSettings["sp.modsite:doclib:examendossier"];
            _edutemplate = ConfigurationManager.AppSettings["sp.edusite:template"];
            _edudoclib_berichten = ConfigurationManager.AppSettings["sp.edusite:doclib:berichten"];
            _edudoclib_examendossier = ConfigurationManager.AppSettings["sp.edusite:doclib:examendossier"];
            _link2mod_list = ConfigurationManager.AppSettings["sp.edusite:list2mod"];
            _link2mod_list_column = ConfigurationManager.AppSettings["sp.edusite:list2mod:column"];
            _link2mod_list_value = ConfigurationManager.AppSettings["sp.edusite:list2mod:value"];
            _eventlogsource = ConfigurationManager.AppSettings["eventlogsource"];

            _mailer = new System.Net.Mail.SmtpClient();
            _mailfrom = ConfigurationManager.AppSettings["smtp:from"];
            _mail2business = ConfigurationManager.AppSettings["smtp.sitecreatednotification:to"];
            _mail2admin = ConfigurationManager.AppSettings["smtp.errornotification:to"];
        }

        private void SendNotification2Business(string subject, string body)
        {
            try
            {
                _mailer.SendAsync(_mailfrom, _mail2business, subject, body, "business");
            }
            catch (Exception e)
            {
                EventLog.WriteEntry(_eventlogsource, e.Message + "\n\nmail subject: " + subject + "\nmail body" + body, EventLogEntryType.Error);
            }
        }

        private void MailCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                this.LogWarning("mail", new OperationReport());
            }
            else
            {
                EventLog.WriteEntry(_eventlogsource, e.Error.Message + "\nMAIL ERROR", EventLogEntryType.Error);
            }
        }

        private void SendNotification2Admin(string subject, string body)
        {
            _mailer.SendCompleted += new System.Net.Mail.SendCompletedEventHandler(MailCompleted);
            try
            {
                _mailer.SendAsync(_mailfrom, _mail2admin, subject, body, "admin");
            }
            catch (Exception e)
            {
                EventLog.WriteEntry(_eventlogsource, e.Message + "\n\nmail subject: " + subject + "\nmail body" + body, EventLogEntryType.Error);
            }
        }




        #region ErrorHandlingAndLogging
        /// <summary>
        /// write msg as info to eventlog and report
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="report"></param>
        private void LogInfo(string msg, OperationReport report)
        {
            EventLog.WriteEntry(_eventlogsource, msg, EventLogEntryType.Information);
            report.Messages.Add(msg);
        }
        /// <summary>
        /// write msg as error to eventlog,
        /// dump report as error to eventlog,
        /// throw FaultException(msg)
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="report"></param>
        private void LogError(string msg, OperationReport report)
        {
            string tracedump = "Operation trace:";
            foreach (string item in report.Messages)
            {
                tracedump += "\n\t";
                tracedump += item;
            }
            EventLog.WriteEntry(_eventlogsource, msg, EventLogEntryType.Error);
            EventLog.WriteEntry(_eventlogsource, tracedump, EventLogEntryType.Error);
            this.SendNotification2Admin("OGN.SharePoint.Services: error", msg + "\n" + tracedump);
            throw new FaultException<string>(msg, "Error");
        }
        /// <summary>
        /// write msg as error to eventlog,
        /// dump report as error to eventlog,
        /// write e.Message as error to eventlog,
        /// throw FaultException(msg)
        /// </summary>
        /// <param name="e"></param>
        /// <param name="msg"></param>
        /// <param name="report"></param>
        private void LogException(Exception e, string msg, OperationReport report)
        {
            string tracedump = "Operation trace:";
            foreach (string item in report.Messages)
            {
                tracedump += "\n\t";
                tracedump += item;
            }
            EventLog.WriteEntry(_eventlogsource, msg, EventLogEntryType.Error);
            EventLog.WriteEntry(_eventlogsource, tracedump, EventLogEntryType.Error);
            EventLog.WriteEntry(_eventlogsource, e.Message, EventLogEntryType.Error);
            this.SendNotification2Admin("OGN.SharePoint.Services: exceptie", msg + "\n" + tracedump + "\n" + e.Message);
            //throw new FaultException<string>(msg, "Exception");

        }
        /// <summary>
        /// write msg as warning to eventlog and report
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="report"></param>
        private void LogWarning(string msg, OperationReport report)
        {
            string tracedump = "Operation trace:";
            foreach (string item in report.Messages)
            {
                tracedump += "\n\t";
                tracedump += item;
            }
            try
            {
                EventLog.WriteEntry(_eventlogsource, msg + "\n\n" + tracedump, EventLogEntryType.Warning);
            }
            catch (Exception ex)
            {
                // No catch
            }
            report.Messages.Add(msg);
            report.ResultType = OperationResultType.Warning;
        }
        #endregion

        #region SharepointFunctions
        /// <summary>
        /// get context for SP Site
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private ClientContext GetSite(string url, IEduModSite edumod)
        {
            // Check each individual site collection for site existance
            if (edumod.GetType().ToString().Contains("EduProgramme"))
            {
                for (int siteCount = 0; siteCount <= _eduSiteCollectionCount; siteCount++)
                {
                    string siteUrl = url + "opleiding" + siteCount.ToString("0#");
                    string fullUrl = url + "opleiding" + siteCount.ToString("0#") + "/" + edumod.Code;
                    ClientContext ctx = new ClientContext(siteUrl);
                    ctx.Credentials = _creds;
                    var web = ctx.Web;
                    ctx.Load(web, w => w.Webs.Where(webs => webs.Url == fullUrl));
                    ctx.Load(web, w => w.Url);
                    ctx.ExecuteQuery();
                    if (web.Webs.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        edumod.Url = web.Url + "/" + edumod.GetSiteName();
                        return ctx;
                    }
                }
            }
            if (edumod.GetType().ToString().Contains("Module"))
            {
                for (int siteCount = 0; siteCount <= _modSiteCollectionCount; siteCount++)
                {
                    string siteUrl = url + "module" + siteCount.ToString("0#");
                    string fullUrl = url + "module" + siteCount.ToString("0#") + "/" + edumod.Code;
                    ClientContext ctx = new ClientContext(siteUrl);
                    ctx.Credentials = _creds;
                    var web = ctx.Web;
                    ctx.Load(web, w => w.Webs.Where(webs => webs.Url == fullUrl));
                    ctx.Load(web, w => w.Url);
                    ctx.ExecuteQuery();
                    if (web.Webs.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        edumod.Url = web.Url + "/" + edumod.GetSiteName();
                        return ctx;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// get context for SP Site
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private ClientContext GetSite(string url)
        {
            ClientContext ctx = new ClientContext(url);
            ctx.Credentials = _creds;
            return ctx;
        }

        /// <summary>
        /// Get the title of the site
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <returns></returns>
        private string GetTitle(ClientContext ctx)
        {
            ctx.Load(ctx.Web);
            ctx.ExecuteQuery();
            return ctx.Web.Title;
        }



        public void DeleteSubsites(string url)
        {
            ClientContext ctx = new ClientContext(url);
            Web site = ctx.Web;
            foreach (Web subsite in site.Webs)
            {
                subsite.DeleteObject();
            }
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// add the title of an eduprogramme or module as a term to SP term store
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod_id">id of term in which to create term</param>
        /// <param name="edumod">eduprogramme or module</param>
        private void AddTerm(ClientContext ctx, Guid edumod_id, IEduModSite edumod)
        {
            TaxonomySession tses = TaxonomySession.GetTaxonomySession(ctx);
            TermStore terms = tses.GetDefaultSiteCollectionTermStore();
            TermGroup loi_group = terms.GetGroup(_loi_id);
            TermSet cat_set = terms.GetTermSet(_cat_id);
            Term edumod_set = cat_set.GetTerm(edumod_id);

            // Check if term exists
            try
            {
                Term term = edumod_set.CreateTerm(edumod.GetTitle(), 1033, Guid.NewGuid());
                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                // Term exists.
            }
        }

        /// <summary>
        /// create a SP site for an eduprogramme or module
        /// </summary>
        /// <returns>full url to site</returns>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod">eduprogramme or module</param>
        /// <param name="template">site template id</param>
        private string CreateSite(IEduModSite edumod, string template)
        {
            string siteUrl = GetAvailableSiteUrl(edumod);
            ClientContext ctx = new ClientContext(siteUrl);
            this.CreateSite(ctx, edumod.GetTitle(), edumod.GetSiteName(), template);
            Web site = ctx.Web;
            ctx.Load(site);
            ctx.ExecuteQuery();
            string fullUrl = ctx.Web.Url + "/" + edumod.GetSiteName();
            return fullUrl;
        }

        private string GetWebTemplateName(ClientContext ctx, string template)
        {
            Web site = ctx.Web;
            WebTemplateCollection templates = ctx.Web.GetAvailableWebTemplates(1043, false);
            ctx.Load(templates);
            ctx.ExecuteQuery();
            // Get site using call
            var templateGet = (from t in templates
                               where t.Title.ToLower() == template.ToLower()
                               select t).FirstOrDefault();
            return templateGet.Name;
        }

        private string GetAvailableSiteUrl(IEduModSite edumod)
        {
            // Check each individual site collection for site existance
            if (edumod.GetType().ToString().Contains("EduProgramme"))
            {
                for (int siteCount = 0; siteCount <= _eduSiteCollectionCount; siteCount++)
                {
                    string siteUrl = _edu_url + "opleiding" + siteCount.ToString("0#");
                    ClientContext ctx = new ClientContext(siteUrl);
                    ctx.Credentials = _creds;
                    var web = ctx.Web;
                    ctx.Load(web, w => w.Webs);
                    ctx.ExecuteQuery();
                    if (web.Webs.Count >= _eduMaxSiteCount)
                    {
                        continue;
                    }
                    else
                    {
                        return siteUrl;
                    }
                }
                return null;
            }

            if (edumod.GetType().ToString().Contains("Module"))
            {
                for (int siteCount = 0; siteCount <= _modSiteCollectionCount; siteCount++)
                {
                    string siteUrl = _mod_url + "module" + siteCount.ToString("0#");
                    ClientContext ctx = new ClientContext(siteUrl);
                    ctx.Credentials = _creds;
                    var web = ctx.Web;
                    ctx.Load(web, w => w.Webs);
                    ctx.ExecuteQuery();
                    if (web.Webs.Count >= _modMaxSiteCount)
                    {
                        continue;
                    }
                    else
                    {
                        return siteUrl;
                    }
                }
                return null;
            }
            return null;
        } // end function

        /// <summary>
        /// create a SP site
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="title">title of site</param>
        /// <param name="url">url of site</param>
        /// <param name="template">site template id</param>
        private void CreateSite(ClientContext ctx, string title, string url, string template)
        {
            Web site = ctx.Web;
            WebCreationInformation newsite = new WebCreationInformation();
            newsite.WebTemplate = GetWebTemplateName(ctx, template);
            newsite.Title = title;
            newsite.Url = url;
            newsite.UseSamePermissionsAsParentSite = true;
            newsite.Language = _lcid;
            site.Webs.Add(newsite);
            ctx.ExecuteQuery();
        }

        private enum SiteType { edu, mod };

        /// <summary>
        /// Changes permission for sites, lists or doclibraries
        /// </summary>
        /// <param name="ctx">Web Context</param>
        /// <param name="sitetype">MOD or EDU</param>
        private void ChangePermissions(ClientContext ctx, SiteType sitetype)
        {
            foreach (ConfigurationSection sect in configfile.Sections)
            {
                string name = sect.SectionInformation.Name;
                if (name.ToLower().StartsWith("sp.sitepermissions."))
                {
                    char[] delim = { '.' };
                    string[] split = name.Split(delim);

                    string compareSiteType = split[2].ToLower();
                    if (compareSiteType == sitetype.ToString().ToLower())
                    {
                        string configType = split[3].ToLower();
                        SitePermissionsSection config = (SitePermissionsSection)ConfigurationManager.GetSection(name);

                        switch (configType)
                        {
                            case "deactivate":
                                {
                                    break; // No change needed. This is handled in other ChangePermissions function
                                }
                            case "doclib":
                                {
                                    string doclib = string.Empty;
                                    doclib = split[4];
                                    Web site = ctx.Web;
                                    List list = site.Lists.GetByTitle(doclib);
                                    list.BreakRoleInheritance(false, false);

                                    ctx.ExecuteQuery();
                                    foreach (PermissionBindingConfigElement item in config.Permissions)
                                    {
                                        Group sitegroup = site.SiteGroups.GetByName(item.SiteGroup);

                                        RoleDefinition permission = site.RoleDefinitions.GetByName(item.Permission);
                                        RoleDefinitionBindingCollection rdbs = new RoleDefinitionBindingCollection(ctx);

                                        rdbs.Add(permission);
                                        list.RoleAssignments.Add(sitegroup, rdbs);
                                    }
                                    ctx.ExecuteQuery();
                                    break;
                                }
                            case "site":
                                {
                                    string siteName = string.Empty;
                                    siteName = split[4].ToLower();
                                    Web site = null;
                                    if (config.PermissionType.ToLower() == "dynamic")
                                    {
                                        site = ctx.Web.Webs.First(w => w.Title.ToLower().Contains(siteName));
                                    }
                                    else
                                    {
                                        site = ctx.Web.Webs.First(w => w.Title.ToLower() == siteName);

                                    }

                                    site.BreakRoleInheritance(false, false);
                                    ctx.ExecuteQuery();
                                    foreach (PermissionBindingConfigElement item in config.Permissions)
                                    {
                                        Group sitegroup = site.SiteGroups.GetByName(item.SiteGroup);

                                        RoleDefinition permission = site.RoleDefinitions.GetByName(item.Permission);
                                        RoleDefinitionBindingCollection rdbs = new RoleDefinitionBindingCollection(ctx);

                                        rdbs.Add(permission);
                                        site.RoleAssignments.Add(sitegroup, rdbs);
                                    }
                                    ctx.ExecuteQuery();
                                    break;
                                }
                        } // end switch
                    } // end check site type
                }  // end check config section check
            } // end for config
        } // end f


        /// <summary>
        /// returns true if site for eduprogramme or module exists
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod">eduprogramme or module</param>
        /// <returns></returns>
        private bool SiteExists(ClientContext ctx, IEduModSite edumod)
        {
            Web site = ctx.Web;
            ctx.Load(site.Webs, sites => sites.Include(subsite => subsite.Url));
            ctx.ExecuteQuery();

            int siteCount = site.Webs.Count(subsite => subsite.Url.EndsWith("/" + edumod.GetSiteName()));
            if (siteCount > 0)
            {
                edumod.Url = ctx.Url + "/" + edumod.GetSiteName();
                return true;
            }
            return false;
        }

        /// <summary>
        /// returns true if site for subsite 
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod">eduprogramme or module</param>
        /// <returns></returns>
        private bool SiteExists(ClientContext ctx, string subSite)
        {
            Web site = ctx.Web;

            ctx.Load(site.Webs);
            ctx.ExecuteQuery();

            int count = site.Webs.Count(subsite => subsite.Url.EndsWith("/" + subSite.ToLower()));

            if (count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// changes the title of the site for a eduprogramme or module
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod">eduprogramme or module</param>
        private bool ChangeTitle(ClientContext ctx, IEduModSite edumod)
        {
            Web site = ctx.Web;
            ctx.Load(site);
            ctx.ExecuteQuery();
            if (site.Title != edumod.GetTitle())
            {
                site.Title = edumod.GetTitle();
                site.Update();
                ctx.ExecuteQuery();
                if (edumod.GetType() == typeof(ModuleVal))
                {
                    Module mod = (Module)edumod;
                    // Change subsite title too
                    string modTitle = _modsub_title + " " + edumod.GetTitle();
                    ctx.Load(site.Webs);
                    ctx.ExecuteQuery();
                    foreach (Web web in site.Webs)
                    {
                        if (web.Title.Contains(_modsub_title))
                        {
                            web.Title = modTitle;
                            web.Update();
                            ctx.ExecuteQuery();
                        }
                    } // end for
                } // end if
                return true;
            }  // end if
            else
            {
                return false;
            }
        }

        /// <summary>
        /// breaks the site permissions inheritance and sets permissions as configured in the web.config
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod">eduprogramme or module</param>
        private void ChangePermissions(ClientContext ctx, bool isEdu)
        {
            Web site = ctx.Web;
            site.BreakRoleInheritance(false, false);
            /* BreakRoleInheritance(copyRoleAssignments,clearSubscopes)
             * copyRoleAssignments
             *   Type: System.Boolean
             *   Specifies whether to copy the role assignments from the parent securable object.
             *   If the value is false, the collection of role assignments must contain only 1 role assignment containing the current user after the operation.
             * clearSubscopes
             *   Type: System.Boolean
             *   If the securable object is a site, and the clearsubscopes parameter is true, the role assignments for all child securable objects in the current site and in the sites which inherit role assignments from the current site must be cleared and those securable objects will inherit role assignments from the current site after this call.
             *   If the securable object is a site, and the clearsubscopes parameter is false, the role assignments for all child securable objects which do not inherit role assignments from their parent object must remain unchanged.
             */
            ctx.ExecuteQuery();

            SitePermissionsSection config;
            if (isEdu)
            {
                config = (SitePermissionsSection)ConfigurationManager.GetSection("sp.sitepermissions.edu.deactivate");
            }
            else
            {
                config = (SitePermissionsSection)ConfigurationManager.GetSection("sp.sitepermissions.mod.deactivate");
            }

            foreach (PermissionBindingConfigElement item in config.Permissions)
            {
                Group sitegroup = site.SiteGroups.GetByName(item.SiteGroup);

                RoleDefinition permission = site.RoleDefinitions.GetByName(item.Permission);
                RoleDefinitionBindingCollection rdbs = new RoleDefinitionBindingCollection(ctx);

                rdbs.Add(permission);
                site.RoleAssignments.Add(sitegroup, rdbs);
            }
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// creates a link to the site of an eduprogramme or module
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links in which a link is created</param>
        /// <param name="linkto">the eduprogramme or module to which the link targets</param>
        private void CreateLink(ClientContext ctx, string listtitle, string linktourl, string linktodescr, string column, string value, string schoolcolumn, string schoolvalue)
        {
            Web site = ctx.Web;
            List list = site.Lists.GetByTitle(listtitle);

            ListItemCreationInformation itemInfo = new ListItemCreationInformation();
            ListItem item = list.AddItem(itemInfo);
            FieldUrlValue url = new FieldUrlValue();
            url.Url = linktourl; //linkto.GetUrl(_home_url);
            url.Description = linktodescr; //linkto.GetTitle();
            item["URL"] = url;
            if (!string.IsNullOrEmpty(column))
            {
                item[column] = value;
            }

            if (!string.IsNullOrEmpty(schoolcolumn))
            {
                item[schoolcolumn] = schoolvalue;
            }

            item.Update();
            ctx.ExecuteQuery();
        }
        private void CreateLink(ClientContext ctx, string listtitle, string linktourl, string linktodescr)
        {
            CreateLink(ctx, listtitle, linktourl, linktodescr, string.Empty, string.Empty, string.Empty, string.Empty);
        }



        /// <summary>
        /// returns true if a link to the site of an eduprogramme or module exists in the list of links
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links</param>
        /// <param name="linkto">the eduprogramme or module to which the link targets</param>
        private bool LinkExists(ClientContext ctx, string listtitle, string linktourl)
        {
            Web site = ctx.Web;
            List list = site.Lists.GetByTitle(listtitle);

            CamlQuery qry = new CamlQuery();
            //qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>announce</Value></Eq></Where></Query></View>";
            ListItemCollection items = list.GetItems(qry);
            ctx.Load(items);
            ctx.ExecuteQuery();

            int itemCount = items.Count(item => ((FieldUrlValue)item["URL"]).Url.Equals(linktourl));
            if (itemCount <= 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// updates the descriptions of all links to the site of an eduprogramme or module.
        /// those sites are all the sites to which this site links
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links</param>
        /// <param name="edumod">the eduprogramme or module</param>
        private bool UpdateAllLinksToEduOrMod(ClientContext ctx, string listtitle, IEduModSite edumod)
        {
            Web site = ctx.Web;

            List list = site.Lists.GetByTitle(listtitle);

            CamlQuery qry = new CamlQuery();
            //qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>announce</Value></Eq></Where></Query></View>";
            ListItemCollection items = list.GetItems(qry);
            ctx.Load(items);
            ctx.ExecuteQuery();
            bool isFound = false;
            foreach (ListItem item in items)
            {
                FieldUrlValue sitelinkstome = (FieldUrlValue)item["URL"];
                ClientContext ctx2 = this.GetSite(sitelinkstome.Url);
                //there are two names for lists of links on an edu or mod sites, mod sites have links to edus, edu sites have links to mods.
                //listtitle is one name, listtitle2 must be the other
                string listtitle2 = (listtitle.Equals(_link2edu_list)) ? _link2mod_list : _link2edu_list;
                this.UpdateLink(ctx2, listtitle2, edumod);
                isFound = true;
            }


            return isFound;
        }

        /// <summary>
        /// updates the description of a link in the list
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links</param>
        /// <param name="linkto">the eduprogramme or module to which the link targets</param>
        private void UpdateLink(ClientContext ctx, string listtitle, IEduModSite linkto)
        {
            UpdateLink(ctx, listtitle, linkto.GetTitle(), linkto.Url, string.Empty, string.Empty, string.Empty, string.Empty);
        }
        private void UpdateLink(ClientContext ctx, string listtitle, string linktitle, string linkurl, string column, string val, string schoolcolumn, string schoolvalue)
        {
            Web site = ctx.Web;
            List list = site.Lists.GetByTitle(listtitle);
            CamlQuery qry = new CamlQuery();
            //qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>announce</Value></Eq></Where></Query></View>";
            ListItemCollection items = list.GetItems(qry);
            ctx.Load(items);
            ctx.ExecuteQuery();
            bool isFound = false;
            foreach (ListItem item in items)
            {
                FieldUrlValue url = (FieldUrlValue)item["URL"];
                if (url.Url.Equals(linkurl))
                {
                    url.Description = linktitle;
                    item["URL"] = url;
                    if (!string.IsNullOrEmpty(column))
                    {
                        item[column] = val;
                    }

                    if (!string.IsNullOrEmpty(schoolcolumn))
                    {
                        item[schoolcolumn] = schoolvalue;
                    }
                    item.Update();
                    isFound = true;
                    break;
                }
            }
            ctx.ExecuteQuery();

            if (isFound == false)
            {
                this.CreateLink(ctx, listtitle, linkurl, linktitle, column, val, new OperationReport());
            }
        }


        private void UpdateLinkName(ClientContext ctx, string listtitle, string linkurl, string newTitle)
        {
            Web site = ctx.Web;

            List list = site.Lists.GetByTitle(listtitle);

            CamlQuery qry = new CamlQuery();
            //qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>announce</Value></Eq></Where></Query></View>";
            ListItemCollection items = list.GetItems(qry);
            ctx.Load(items);
            ctx.ExecuteQuery();

            foreach (ListItem item in items)
            {
                FieldUrlValue url = (FieldUrlValue)item["URL"];
                if (url.Url.Equals(linkurl))
                {
                    if (newTitle.Equals(string.Empty)) { } else { item["Title"] = newTitle; }
                    item.Update();
                    break;
                }
            }
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// deletes a link in the list
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links</param>
        /// <param name="linkto">the eduprogramme or module to which the link targets</param>
        private void DeleteLink(ClientContext ctx, string listtitle, string linktourl)
        {
            Web site = ctx.Web;

            List list = site.Lists.GetByTitle(listtitle);

            CamlQuery qry = new CamlQuery();
            //qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>announce</Value></Eq></Where></Query></View>";
            ListItemCollection items = list.GetItems(qry);
            ctx.Load(items);
            ctx.ExecuteQuery();

            foreach (ListItem item in items)
            {
                FieldUrlValue url = (FieldUrlValue)item["URL"];
                if (url.Url.Equals(linktourl))
                {
                    item.DeleteObject();
                    break;
                }
            }
            ctx.ExecuteQuery();
        }

        private void CreateLink(ClientContext site, string name_link_list, string link, string link_descr, string column, string value, OperationReport report)
        {
            if (this.LinkExists(site, name_link_list, link))
            {
                UpdateLinkName(site, name_link_list, link, link_descr);
                this.LogWarning("Link niet gemaakt. Link bestaat al. Link name mogelijk aangepast", report);
            }
            else
            {
                this.CreateLink(site, name_link_list, link, link_descr, column, value, string.Empty, string.Empty);
                report.Messages.Add("Link gemaakt.");
            }
        }

        #endregion

        #region ServiceOperations
        /* The operation methods have a pattern:
         * SNIPPET FOR OPERATIONS:
            OperationReport report = new OperationReport();
            try
            {
                //operation here
                //That is, do SharePoint stuff (this.SPFunction(...))
                //report on what has been done (report.Message.Add(msg) or LogInfo(msg) or LogWarning(msg))
                //do more Sharepoint stuff and reporting
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie ???():\n" + e.Message, report); }
            return report;
         */
        public OperationReport Create(EduProgrammeVal edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak opleiding: id->" + edu.Id + ", code->" + edu.Code + ", naam->" + edu.Name + ", LOI site->" + edu.EduWorkSpace);
                ClientContext ctx = this.GetSite(_edu_url, edu);
                if (ctx != null)
                {
                    this.LogWarning("Site niet gemaakt. Opleiding bestaat al.", report);
                }
                else
                {
                    //create the site
                    edu.Url = CreateSite(edu, _edutemplate);
                    report.Messages.Add("Site gemaakt.");
                    ClientContext ctx_edu = this.GetSite(edu.Url);

                    //change permissions on lists, sites and doclibs as configured                          ,
                    ChangePermissions(ctx_edu, SiteType.edu);
                    report.Messages.Add("Permissies van Doc.Libs en lijsten op site aangepast.");

                    // Opleidingsmatrix
                    ctx = this.GetSite(_eduHome_url);
                    CreateLink(ctx, _edu_siteslist, edu.Url, edu.Url, _edu_siteslist_column, edu.GetTitle(), _edu_siteslist_school_column, edu.EduType);
                    report.Messages.Add("Link vanaf sitecollectie naar opleidingssite gemaakt.");

                    ctx = this.GetSite(edu.Url);
                    if (!string.IsNullOrEmpty(edu.EduWorkSpace))
                    {
                        report.Messages.Add("Check voor link naar andere opleiding");
                        EduProgramme loiEdu = new EduProgramme();
                        loiEdu.Code = edu.EduWorkSpace;
                        EduProgrammeRef loisite = new EduProgrammeRef(edu.EduWorkSpace, loiEdu);
                        ClientContext ctx_check = this.GetSite(_edu_url, loisite);
                        ClientContext ctx_loi = this.GetSite(loisite.Url);
                        if (ctx_loi != null)
                        {
                            if (LinkExists(ctx, _link2mod_list, loisite.Url) == false)
                            {
                                CreateLink(ctx, _link2mod_list, loisite.Url, this.GetTitle(ctx_loi), _link2edu_list_column, "LOI " + _link2edu_list_value, report);
                                report.Messages.Add("Link naar LOI opleidingssite gemaakt.");
                            }
                            else
                            {
                                report.Messages.Add("Link naar LOI opleidingssite was er al.");
                            }
                            if (LinkExists(ctx_loi, _link2mod_list, edu.Url) == false)
                            {
                                CreateLink(ctx_loi, _link2mod_list, edu.Url, edu.GetTitle(), _link2edu_list_column, edu.EduType + " " + _link2edu_list_value, report);
                                report.Messages.Add("Link vanuit LOI opleidingssite gemaakt.");
                            }
                            else
                            {
                                report.Messages.Add("Link vanuit LOI opleidingssite was er al.");
                            }
                        }
                        else
                        {
                            this.LogWarning("Check: Link vanuit LOI opleidingssite niet gemaakt. LOI opleidingssite bestaat niet.", report);
                        }
                    }
                    this.AddTerm(ctx, _edu_id, edu);
                    report.Messages.Add("Term gemaakt.");
                    this.SendNotification2Business("Nieuwe SharePoint site voor opleiding '" + edu.GetTitle() + "'", edu.Url);
                }
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(edu):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Update(EduProgrammeVal edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Wijzig opleidingsnaam: id->" + edu.Id + ", code->" + edu.Code + ", nieuwe naam->" + edu.Name);

                ClientContext ctx = this.GetSite(edu.Url);
                if (this.ChangeTitle(ctx, edu))
                {
                    report.Messages.Add("Site titel gewijzigd.");
                }
                else
                {
                    report.Messages.Add("Site titel hoefde niet gewijzigd te worden.");
                }

                if (this.UpdateAllLinksToEduOrMod(ctx, _link2mod_list, edu))
                {
                    report.Messages.Add("Beschrijvingen van links naar deze site geupdate.");
                }
                else
                {
                    report.Messages.Add("Beschrijvingen van links naar deze site hoefden niet geupdate te worden.");
                }

                ClientContext ctx_home = this.GetSite(_eduHome_url);
                UpdateLink(ctx_home, _edu_siteslist, edu.Url, edu.Url, _edu_siteslist_column, edu.GetTitle(), _edu_siteslist_school_column, edu.EduType);
                report.Messages.Add("Link vanaf sitecollectie naar opleidingssite aangepast.");
                this.AddTerm(ctx, _edu_id, edu);
                report.Messages.Add("Term gemaakt.");

                if (!string.IsNullOrEmpty(edu.EduWorkSpace))
                {
                    report.Messages.Add("Check voor link naar andere opleiding");
                    EduProgramme loiEdu = new EduProgramme();
                    loiEdu.Code = edu.EduWorkSpace;
                    EduProgrammeRef loisite = new EduProgrammeRef(edu.EduWorkSpace, loiEdu);
                    ClientContext ctx_check = this.GetSite(_edu_url, loisite);
                    ClientContext ctx_loi = this.GetSite(loisite.Url);
                    if (ctx_loi != null)
                    {
                        if (LinkExists(ctx, _link2mod_list, loisite.Url) == false)
                        {
                            CreateLink(ctx, _link2mod_list, loisite.Url, this.GetTitle(ctx_loi), _link2edu_list_column, "LOI " + _link2edu_list_value, report);
                            report.Messages.Add("Link naar LOI opleidingssite gemaakt.");
                        }
                        else
                        {
                            report.Messages.Add("Link naar LOI opleidingssite was er al.");
                        }
                        if (LinkExists(ctx_loi, _link2mod_list, edu.Url) == false)
                        {
                            CreateLink(ctx_loi, _link2mod_list, edu.Url, edu.GetTitle(), _link2edu_list_column, edu.EduType + " " + _link2edu_list_value, report);
                            report.Messages.Add("Link vanuit LOI opleidingssite gemaakt.");
                        }
                        else
                        {
                            report.Messages.Add("Link vanuit LOI opleidingssite was er al.");
                        }
                    }
                    else
                    {
                        this.LogWarning("Check: Link vanuit LOI opleidingssite niet gemaakt. LOI opleidingssite bestaat niet.", report);
                    }
                }
            }
            catch (Exception e)
            {
                this.LogException(e, "Er is een fout opgetreden tijdens operatie Update(edu):\n" + e.Message, report);
            }
            return report;
        }

        public OperationReport Delete(EduProgrammeRef edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Deactiveer opleidingssite: id->" + edu.Id);
                ClientContext ctx = this.GetSite(edu.Url);
                this.ChangePermissions(ctx, true);
                report.Messages.Add("Permissies ingetrokken.");
                this.SendNotification2Business("Permissies gewijzigd van SharePoint site voor opleiding"
                               , "De permissies zijn gewijzigd omdat de opleiding inactief is geworden.\n" + edu.Url);
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Delete(edu):\n" + e.Message, report); }
            return report;
        }


        public OperationReport Create(ModuleVal mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak module: id->" + mod.Id + ", code->" + mod.Code + ", naam->" + mod.Name);
                ClientContext ctx = this.GetSite(_mod_url, mod);
                if (ctx != null)
                {
                    this.LogWarning("Site niet gemaakt. Module bestaat al.", report);
                    //create subsite, add module name to subsite
                    ClientContext ctx_mod = this.GetSite(mod.Url);
                    if (this.SiteExists(ctx_mod, _modsub_title))
                    {
                        this.LogWarning("SubSite niet gemaakt. Site bestaat al.", report);
                    }
                    else
                    {
                        Console.WriteLine("Subsite bestond nog niet");
                        string modTitle = _modsub_title + " " + mod.GetTitle();
                        this.CreateSite(ctx_mod, modTitle, _modsub_id, _modsubtemplate);
                        report.Messages.Add("Subsite gemaakt.");

                        //change permissions on lists, sites and doclibs as configured
                        ChangePermissions(ctx_mod, SiteType.mod);
                        report.Messages.Add("Permissies van Doc.Libs en lijsten op site aangepast.");
                    }
                }
                else
                {
                    //create module site
                    mod.Url = this.CreateSite(mod, _modtemplate);
                    report.Messages.Add("Site gemaakt.");

                    //create subsite, add module name to subsite
                    ClientContext ctx_mod = this.GetSite(mod.Url);
                    string modTitle = _modsub_title + " " + mod.GetTitle();
                    this.CreateSite(ctx_mod, modTitle, _modsub_id, _modsubtemplate);
                    report.Messages.Add("Subsite gemaakt.");

                    //change permissions on lists, sites and doclibs as configured
                    ChangePermissions(ctx_mod, SiteType.mod);
                    report.Messages.Add("Permissies van Doc.Libs en lijsten op site aangepast.");

                    // Modulematrix
                    ctx = this.GetSite(_modHome_url);
                    CreateLink(ctx, _mod_siteslist, mod.Url, mod.Url, _mod_siteslist_column, mod.GetTitle(), _mod_siteslist_school_column, "LOI");
                    report.Messages.Add("Link vanaf sitecollectie naar modulesite gemaakt.");

                    ctx = this.GetSite(mod.Url);
                    //create links from and to module site
                    if (!string.IsNullOrEmpty(mod.LinkedModule))
                    {
                        report.Messages.Add("Maak link van Studieplan naar LOI");
                        Module loiMod = new Module();
                        loiMod.Code = mod.LinkedModule;
                        ModuleRef loisite = new ModuleRef(mod.LinkedModule, loiMod);
                        ClientContext ctx_check = this.GetSite(_mod_url, loisite);
                        ClientContext ctx_loi = this.GetSite(loisite.Url);
                        if (ctx_loi != null)
                        {
                            if (LinkExists(ctx, _link2mod_list, loisite.Url) == false)
                            {
                                CreateLink(ctx, _link2edu_list, loisite.Url, this.GetTitle(ctx_loi), _link2mod_list_column, "LOI " + _link2mod_list_value, report);
                                report.Messages.Add("Link van Studieplan site naar LOI modulesite gemaakt.");
                            }
                            else
                            {
                                report.Messages.Add("Link naar LOI opleidingssite was er al.");
                            }
                            if (LinkExists(ctx_loi, _link2mod_list, mod.Url) == false)
                            {
                                CreateLink(ctx_loi, _link2edu_list, mod.Url, mod.GetTitle(), _link2mod_list_column, "Studieplan " + _link2mod_list_value, report);
                                report.Messages.Add("Link vanuit LOI modulesite naar Studieplan site gemaakt.");
                            }
                            else
                            {
                                report.Messages.Add("Link vanuit LOI opleidingssite was er al.");
                            }
                        }
                        else
                        {
                            this.LogWarning("Link vanuit LOI modulesite niet gemaakt. LOI modulesite bestaat niet.", report);
                        }

                        //create links from opleiding and to module site
                        if (!string.IsNullOrEmpty(mod.EduCode))
                        {
                            report.Messages.Add("Maak link van Modulesite naar Opleidingssite");
                            Link link = new Link();
                            ModuleRef modRef = new ModuleRef(mod.Code, mod);
                            modRef.Url = mod.Url;
                            EduProgramme edu = new EduProgramme();
                            edu.Id = mod.EduCode;
                            edu.Code = mod.EduCode;
                            EduProgrammeRef eduRef = new EduProgrammeRef(mod.EduCode, edu);
                            ClientContext ctx_edu = GetSite(_edu_url, edu);
                            eduRef.Url = edu.Url;
                            link.Module = modRef;
                            link.EduProgramme = eduRef;
                            Create(link);
                        }

                        //create term in store
                        this.AddTerm(ctx, _mod_id, mod);
                        report.Messages.Add("Term gemaakt.");
                        this.SendNotification2Business("Nieuwe SharePoint site voor module '" + mod.GetTitle() + "'", mod.Url);
                    }
                }
            }

            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(mod):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Update(ModuleVal mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Wijzig modulenaam: id->" + mod.Id + ", code->" + mod.Code);
                ClientContext ctx = this.GetSite(mod.Url);
                if (this.ChangeTitle(ctx, mod))
                {
                    report.Messages.Add("Site titel gewijzigd.");
                }
                else
                {
                    report.Messages.Add("Site titel hoefde niet gewijzigd te worden.");
                }
                // Check subsite
                if (this.SiteExists(ctx, _modsub_title))
                {
                    this.LogWarning("SubSite niet gemaakt. Site bestaat al.", report);
                }
                else
                {
                    Console.WriteLine("Subsite bestond nog niet");
                    string modTitle = _modsub_title + " " + mod.GetTitle();
                    this.CreateSite(ctx, modTitle, _modsub_id, _modsubtemplate);
                    report.Messages.Add("Subsite gemaakt.");

                    //change permissions on lists, sites and doclibs as configured
                    ChangePermissions(ctx, SiteType.mod);
                    report.Messages.Add("Permissies van Doc.Libs en lijsten op site aangepast.");
                }

                if (this.UpdateAllLinksToEduOrMod(ctx, _link2mod_list, mod))
                {
                    report.Messages.Add("Beschrijvingen van links naar deze site geupdate.");
                }
                else
                {
                    report.Messages.Add("Beschrijvingen van links naar deze site hoefden niet geupdate te worden.");
                }

                ClientContext ctx_home = this.GetSite(_modHome_url);
                UpdateLink(ctx_home, _mod_siteslist, mod.Url, mod.Url, _mod_siteslist_column, mod.GetTitle(), _mod_siteslist_school_column, "LOI");
                report.Messages.Add("Link vanaf sitecollectie naar modulesite aangepast.");

                //create links from and to module site
                if (!string.IsNullOrEmpty(mod.LinkedModule))
                {
                    report.Messages.Add("Check link van Studieplan naar LOI");
                    Module loiMod = new Module();
                    loiMod.Code = mod.LinkedModule;
                    ModuleRef loisite = new ModuleRef(mod.LinkedModule, loiMod);
                    ClientContext ctx_check = this.GetSite(_mod_url, loisite);
                    ClientContext ctx_loi = this.GetSite(loisite.Url);
                    if (ctx_loi != null)
                    {
                        if (LinkExists(ctx, _link2edu_list, loisite.Url) == false)
                        {
                            CreateLink(ctx, _link2edu_list, loisite.Url, this.GetTitle(ctx_loi), _link2mod_list_column, "LOI " + _link2mod_list_value, report);
                            report.Messages.Add("Link van Studieplan site naar LOI modulesite gemaakt.");
                        }
                        else
                        {
                            report.Messages.Add("Link van Studieplan site naar LOI modulesite was er al.");
                        }
                        if (LinkExists(ctx_loi, _link2edu_list, loisite.Url) == false)
                        {
                            CreateLink(ctx_loi, _link2edu_list, mod.Url, mod.GetTitle(), _link2mod_list_column, "Studieplan " + _link2mod_list_value, report);
                            report.Messages.Add("Link vanuit LOI modulesite naar Studieplan site gemaakt.");
                        }
                        else
                        {
                            report.Messages.Add("Link vanuit LOI modulesite naar Studieplan site was er al.");
                        }
                    }
                    else
                    {
                        this.LogWarning("Link vanuit LOI modulesite niet gemaakt. LOI modulesite bestaat niet.", report);
                    }
                }

                //create links from opleiding and to module site
                if (!string.IsNullOrEmpty(mod.EduCode))
                {
                    report.Messages.Add("Check link van Modulesite naar Opleidingssite");
                    Link link = new Link();
                    ModuleRef modRef = new ModuleRef(mod.Code, mod);
                    modRef.Url = mod.Url;
                    EduProgramme edu = new EduProgramme();
                    edu.Id = mod.EduCode;
                    edu.Code = mod.EduCode;
                    EduProgrammeRef eduRef = new EduProgrammeRef(mod.EduCode, edu);
                    ClientContext ctx_edu = GetSite(_edu_url, edu);
                    eduRef.Url = edu.Url;
                    link.Module = modRef;
                    link.EduProgramme = eduRef;
                    Create(link);
                }
                // TODO: Rename ontwikkeldossier

                this.AddTerm(ctx, _mod_id, mod);
                report.Messages.Add("Term gemaakt.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Update(mod):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Delete(ModuleRef mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Deactiveer modulesite: id->" + mod.Id);
                ClientContext ctx = this.GetSite(mod.Url);
                this.ChangePermissions(ctx, false);
                report.Messages.Add("Permissies ingetrokken.");
                this.SendNotification2Business("Permissies gewijzigd van SharePoint site voor module"
                               , "De permissies zijn gewijzigd omdat de module inactief is geworden.\n" + mod.Url);
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Delete(mod):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Create(Link link)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak links: opl_id->" + link.EduProgramme.Id + ", mod_id->" + link.Module.Id);
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.Url);
                ClientContext ctx_mod = this.GetSite(link.Module.Url);

                if (this.LinkExists(ctx_edu, _link2mod_list, link.Module.Url))
                {
                    Console.WriteLine("Link naar modulesite niet gemaakt. Link bestaat al.");
                    this.LogWarning("Link naar modulesite niet gemaakt. Link bestaat al.", report);
                }
                else
                {
                    this.CreateLink(ctx_edu, _link2mod_list, link.Module.Url, this.GetTitle(ctx_mod), _link2mod_list_column, _link2mod_list_value, string.Empty, string.Empty);
                    report.Messages.Add("Link naar modulesite gemaakt.");
                }
                if (this.LinkExists(ctx_mod, _link2edu_list, link.EduProgramme.Url))
                {
                    Console.WriteLine("Link naar opleidingssite niet gemaakt. Link bestaat al.");
                    this.LogWarning("Link naar opleidingssite niet gemaakt. Link bestaat al.", report);
                }
                else
                {
                    this.CreateLink(ctx_mod, _link2edu_list, link.EduProgramme.Url, this.GetTitle(ctx_edu), _link2edu_list_column, _link2edu_list_value, string.Empty, string.Empty);
                    report.Messages.Add("Link naar opleidingssite gemaakt.");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e.Message);
                this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(link):\n" + e.Message, report);
            }
            return report;
        }

        public OperationReport Update(UpdateType<Link> change)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Vervang link.");
                OperationReport report1 = this.Create(change.To);
                foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                report.ResultType = report1.ResultType;
                OperationReport report2 = this.Delete(change.From);
                foreach (string msg in report2.Messages) { report.Messages.Add(msg); }
                report.ResultType = (report2.ResultType == OperationResultType.Warning) ? report2.ResultType : report.ResultType;
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Update(link):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Delete(Link link)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Verwijder links: opl_id->" + link.EduProgramme.Id + ", mod_id->" + link.Module.Id);
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.Url);
                this.DeleteLink(ctx_edu, _link2mod_list, link.Module.Url);
                report.Messages.Add("Link naar modulesite verwijderd.");
                ClientContext ctx_mod = this.GetSite(link.Module.Url);
                this.DeleteLink(ctx_mod, _link2edu_list, link.EduProgramme.Url);
                report.Messages.Add("Link naar opleidingssite verwijderd.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Delete(link):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Test()
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Dit is de Test operatie van deze service..");
                this.LogInfo("Er is getest: Test().", report);
                report.Messages.Add("Hier is niet veel gebeurd.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Test():\n" + e.Message, report); }
            return report;
        }

        public OperationReport TestException()
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Dit is de TestFout operatie van deze service.");
                this.LogInfo("Er is getest: TestException().", report);
                throw new ApplicationException("Deze fout is de bedoeling van TestException().");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie TestException():\n" + e.Message, report); }
            return report;
        }
        #endregion



        public OperationReport DoUndeterminedAction(EduProgrammeVal edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Onbepaalde actie op opleiding: id->" + edu.Id + ", code->" + edu.Code + ", naam->" + edu.Name + ", type->" + edu.EduType);
                ClientContext ctx = this.GetSite(_edu_url, edu);
                if (ctx != null)
                {
                    report.Messages.Add("De opleidingssite bestaat al.");
                    ClientContext ctx_edu = this.GetSite(edu.Url);
                    string oldName = this.GetTitle(ctx_edu);
                    string newName = edu.GetTitle();
                    if (oldName.Equals(newName))
                    {
                        report.Messages.Add("De naam van de opleiding hoeft niet gewijzigd te worden.");
                        OperationReport report1 = this.Update(edu);
                        foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                        report.ResultType = report1.ResultType;
                    }
                    else
                    {
                        report.Messages.Add("Actie bepaald: De naam van de opleiding moet gewijzigd worden van '" + oldName + "' naar '" + newName + "'.");
                        OperationReport report1 = this.Update(edu);
                        foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                        report.ResultType = report1.ResultType;
                    }
                }
                else
                {
                    report.Messages.Add("Actie bepaald: De opleidingssite moet gemaakt worden.");
                    OperationReport report1 = this.Create(edu);
                    foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                    report.ResultType = report1.ResultType;
                }
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie DoUndeterminedAction(edu):\n" + e.Message, report); }
            return report;
        }

        public OperationReport DoUndeterminedAction(ModuleVal mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Onbepaalde actie op module: id->" + mod.Id + ", code->" + mod.Code + ", naam->" + mod.Name);
                ClientContext ctx = this.GetSite(_mod_url, mod);
                if (ctx != null)
                {
                    report.Messages.Add("De modulesite bestaat al.");
                    ClientContext ctx_mod = this.GetSite(mod.Url);
                    string oldName = this.GetTitle(ctx_mod);
                    string newName = mod.GetTitle();
                    if (oldName.Equals(newName))
                    {
                        report.Messages.Add("De naam van de module hoeft niet gewijzigd te worden.");
                        OperationReport report1 = this.Update(mod);
                        foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                        report.ResultType = report1.ResultType;
                    }
                    else
                    {
                        report.Messages.Add("Actie bepaald: De naam van de module moet gewijzigd worden van '" + oldName + "' naar '" + newName + "'.");
                        OperationReport report1 = this.Update(mod);
                        foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                        report.ResultType = report1.ResultType;
                    }
                }
                else
                {
                    report.Messages.Add("Actie bepaald: De modulesite moet gemaakt worden.");
                    OperationReport report1 = this.Create(mod);
                    foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                    report.ResultType = report1.ResultType;
                }
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie DoUndeterminedAction(mod):\n" + e.Message, report); }
            return report;
        }

        public void FixSiteNames()
        {
            ClientContext ctx = new ClientContext("https://teamwise.ogn.eu/module");
            ctx.Credentials = _creds;
            ctx.Load(ctx.Web);
            ctx.ExecuteQuery();

            ctx.Load(ctx.Web.Webs, sites => sites.Include(subsite => subsite.Title));
            //var site = ctx.Web.Webs.First(s => s.Title.Contains(';'));


            ctx.ExecuteQuery();

            foreach (Web web in ctx.Web.Webs)
            {
                Console.WriteLine(web.Title);
                ctx.Load(web.Webs, w => w.Include(ont => ont.Title));
                ctx.ExecuteQuery();

                Web subWeb = web.Webs[0];
                if (subWeb.Title.Contains(';') | (subWeb.Title.Contains('"')))
                {
                    Console.WriteLine(subWeb.Title);
                    subWeb.Title = subWeb.Title.Replace("\"", "").Replace(";", "").Trim();
                    Console.WriteLine(subWeb.Title);
                    subWeb.Update();
                    ctx.ExecuteQuery();
                }

            }
        } // end f


    } // end c
} // end ns
