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
            _creds = CredentialCache.DefaultNetworkCredentials;
            //_creds = new NetworkCredential("user", "pass", "ad"); //for testing

            _mod_url = ConfigurationManager.AppSettings["sp.sitecollection:mod:url"];
            _edu_url = ConfigurationManager.AppSettings["sp.sitecollection:edu:url"];

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
            throw new FaultException<string>(msg, "Exception");
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

            Term term = edumod_set.CreateTerm(edumod.GetTitle(), 1033, Guid.NewGuid());
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// create a SP site for an eduprogramme or module
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="edumod">eduprogramme or module</param>
        /// <param name="template">site template id</param>
        private void CreateSite(ClientContext ctx, IEduModSite edumod, string template)
        {
            this.CreateSite(ctx, edumod.GetTitle(), edumod.GetSiteName(), template);
        }

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
            newsite.WebTemplate = template;
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

            return 0 < site.Webs.Count(subsite => subsite.Url.EndsWith("/" + edumod.GetSiteName()));
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
        private void ChangeTitle(ClientContext ctx, IEduModSite edumod)
        {
            Web site = ctx.Web;
            site.Title = edumod.GetTitle();
            site.Update();
            ctx.ExecuteQuery();
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

            return 0 < items.Count(item => ((FieldUrlValue)item["URL"]).Url.Equals(linktourl));
        }

        /// <summary>
        /// updates the descriptions of all links to the site of an eduprogramme or module.
        /// those sites are all the sites to which this site links
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links</param>
        /// <param name="edumod">the eduprogramme or module</param>
        private void UpdateAllLinksToEduOrMod(ClientContext ctx, string listtitle, IEduModSite edumod)
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
                FieldUrlValue sitelinkstome = (FieldUrlValue)item["URL"];
                ClientContext ctx2 = this.GetSite(sitelinkstome.Url);
                //there are two names for lists of links on an edu or mod sites, mod sites have links to edus, edu sites have links to mods.
                //listtitle is one name, listtitle2 must be the other
                string listtitle2 = (listtitle.Equals(_link2edu_list)) ? _link2mod_list : _link2edu_list;
                this.UpdateLink(ctx2, listtitle2, edumod);
            }
        }

        /// <summary>
        /// updates the description of a link in the list
        /// </summary>
        /// <param name="ctx">SP context</param>
        /// <param name="listtitle">the name of the list of links</param>
        /// <param name="linkto">the eduprogramme or module to which the link targets</param>
        private void UpdateLink(ClientContext ctx, string listtitle, IEduModSite linkto)
        {
            UpdateLink(ctx, listtitle, linkto.GetTitle(), linkto.GetUrl(), string.Empty, string.Empty);
        }
        private void UpdateLink(ClientContext ctx, string listtitle, string linktitle, string linkurl, string column, string val)
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
                    url.Description = linktitle;
                    item["URL"] = url;
                    if (column.Equals(string.Empty)) { } else { item[column] = val; }
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
                this.LogWarning("Link niet gemaakt. Link bestaat al.", report);
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
                report.Messages.Add("Maak opleiding: id->" + edu.Id + ", code->" + edu.Code + ", naam->" + edu.Name + ", LOI site->" + edu.LOISite);
                ClientContext ctx = this.GetSite(_edu_url);
                if (this.SiteExists(ctx, edu))
                {
                    this.LogWarning("Site niet gemaakt. Opleiding bestaat al.", report);
                }
                else
                {
                    //create the site
                    this.CreateSite(ctx, edu, _edutemplate);
                    report.Messages.Add("Site gemaakt.");
                    ClientContext ctx_edu = this.GetSite(edu.GetUrl());

                    //change permissions on lists, sites and doclibs as configured
                    ChangePermissions(ctx_edu, SiteType.edu);
                    report.Messages.Add("Permissies van Doc.Libs en lijsten op site aangepast.");

                    // Opleidingsmatrix
                    CreateLink(ctx, _edu_siteslist, edu.GetUrl(), edu.GetUrl(), _edu_siteslist_column, edu.GetTitle(), _edu_siteslist_school_column, edu.EduType);
                    report.Messages.Add("Link vanaf sitecollectie naar opleidingssite gemaakt.");

                    if (!string.IsNullOrEmpty(edu.LOISite))
                    {
                        report.Messages.Add("Maak link van en naar LOI opleidingssite.");
                        EduProgrammeRef loisite = new EduProgrammeRef(edu.LOISite, edu);
                        if (this.SiteExists(ctx, loisite))
                        {
                            ClientContext ctx_loi = this.GetSite(loisite.GetUrl());
                            CreateLink(ctx_edu, _link2mod_list, loisite.GetUrl(), this.GetTitle(ctx_loi), _link2edu_list_column, "LOI " + _link2edu_list_value, report);
                            report.Messages.Add("Link naar LOI opleidingssite gemaakt.");

                            CreateLink(ctx_loi, _link2mod_list, edu.GetUrl(), edu.GetTitle(), _link2edu_list_column, edu.EduType + " " + _link2edu_list_value, report);
                            report.Messages.Add("Link vanuit LOI opleidingssite gemaakt.");
                        }
                        else
                        {
                            this.LogWarning("Link vanuit LOI opleidingssite niet gemaakt. LOI opleidingssite bestaat niet.", report);
                        }
                    }
                    this.AddTerm(ctx, _edu_id, edu);
                    report.Messages.Add("Term gemaakt.");
                    this.SendNotification2Business("Nieuwe SharePoint site voor opleiding '" + edu.GetTitle() + "'", edu.GetUrl());
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

                ClientContext ctx = this.GetSite(edu.GetUrl());
                this.ChangeTitle(ctx, edu);
                report.Messages.Add("Site titel gewijzigd.");
                this.UpdateAllLinksToEduOrMod(ctx, _link2mod_list, edu);
                report.Messages.Add("Beschrijvingen van links naar deze site gewijzigd.");
                ClientContext ctx_home = this.GetSite(_edu_url);
                UpdateLink(ctx_home, _edu_siteslist, edu.GetUrl(), edu.GetUrl(), _edu_siteslist_column, edu.GetTitle());
                report.Messages.Add("Link vanaf sitecollectie naar opleidingssite gemaakt.");
                this.AddTerm(ctx, _edu_id, edu);
                report.Messages.Add("Term gemaakt.");

            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Update(edu):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Delete(EduProgrammeRef edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Deactiveer opleidingssite: id->" + edu.Id);
                ClientContext ctx = this.GetSite(edu.GetUrl());
                this.ChangePermissions(ctx, true);
                report.Messages.Add("Permissies ingetrokken.");
                this.SendNotification2Business("Permissies gewijzigd van SharePoint site voor opleiding"
                               , "De permissies zijn gewijzigd omdat de opleiding inactief is geworden.\n" + edu.GetUrl());
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Delete(edu):\n" + e.Message, report); }
            return report;
        }


        public OperationReport Create(ModuleVal mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak module: id->" + mod.Id + ", code->" + mod.Code + ", naam->" + mod.Name + ", LOI site->" + mod.LOISite);
                ClientContext ctx = this.GetSite(_mod_url);
                if (this.SiteExists(ctx, mod))
                {
                    this.LogWarning("Site niet gemaakt. Module bestaat al.", report);
                    //create subsite, add module name to subsite
                    ClientContext ctx_mod = this.GetSite(mod.GetUrl());
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
                    this.CreateSite(ctx, mod, _modtemplate);
                    report.Messages.Add("Site gemaakt.");

                    //create subsite, add module name to subsite
                    ClientContext ctx_mod = this.GetSite(mod.GetUrl());
                    string modTitle = _modsub_title + " " + mod.GetTitle();
                    this.CreateSite(ctx_mod, modTitle, _modsub_id, _modsubtemplate);
                    report.Messages.Add("Subsite gemaakt.");


                    //change permissions on lists, sites and doclibs as configured
                    ChangePermissions(ctx_mod, SiteType.mod);
                    report.Messages.Add("Permissies van Doc.Libs en lijsten op site aangepast.");

                    // Modulematrix
                    CreateLink(ctx, _mod_siteslist, mod.GetUrl(), mod.GetUrl(), _mod_siteslist_column, mod.GetTitle(), _mod_siteslist_school_column, mod.EduType);
                    report.Messages.Add("Link vanaf sitecollectie naar modulesite gemaakt.");

                    //create links from and to module site
                    if (!string.IsNullOrEmpty(mod.LOISite))
                    {
                        report.Messages.Add("Maak link van en naar LOI modulesite.");
                        ModuleRef loisite = new ModuleRef(mod.LOISite, mod);
                        if (this.SiteExists(ctx, loisite))
                        {
                            ClientContext ctx_loi = this.GetSite(loisite.GetUrl());
                            CreateLink(ctx_mod, _link2edu_list, loisite.GetUrl(), this.GetTitle(ctx_loi), _link2mod_list_column, "LOI " + _link2mod_list_value, report);
                            report.Messages.Add("Link naar LOI modulesite gemaakt.");

                            CreateLink(ctx_loi, _link2edu_list, mod.GetUrl(), mod.GetTitle(), _link2mod_list_column, mod.EduType + " " + _link2mod_list_value, report);
                            report.Messages.Add("Link vanuit LOI modulesite gemaakt.");
                        }
                        else
                        {
                            this.LogWarning("Link vanuit LOI modulesite niet gemaakt. LOI modulesite bestaat niet.", report);
                        }
                    }
                    //create term in store
                    this.AddTerm(ctx, _mod_id, mod);
                    report.Messages.Add("Term gemaakt.");
                    this.SendNotification2Business("Nieuwe SharePoint site voor module '" + mod.GetTitle() + "'", mod.GetUrl());

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
                report.Messages.Add("Wijzig modulenaam: id->" + mod.Id + ", code->" + mod.Code + ", nieuwe naam->" + mod.Name);

                ClientContext ctx = this.GetSite(mod.GetUrl());
                this.ChangeTitle(ctx, mod);
                report.Messages.Add("Site titel gewijzigd.");
                this.UpdateAllLinksToEduOrMod(ctx, _link2edu_list, mod);
                report.Messages.Add("Beschrijvingen van links naar deze site gewijzigd.");
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
                ClientContext ctx = this.GetSite(mod.GetUrl());
                this.ChangePermissions(ctx, false);
                report.Messages.Add("Permissies ingetrokken.");
                this.SendNotification2Business("Permissies gewijzigd van SharePoint site voor module"
                               , "De permissies zijn gewijzigd omdat de module inactief is geworden.\n" + mod.GetUrl());
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
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.GetUrl());
                ClientContext ctx_mod = this.GetSite(link.Module.GetUrl());

                if (this.LinkExists(ctx_edu, _link2mod_list, link.Module.GetUrl()))
                {
                    this.LogWarning("Link naar modulesite niet gemaakt. Link bestaat al.", report);
                }
                else
                {
                    this.CreateLink(ctx_edu, _link2mod_list, link.Module.GetUrl(), this.GetTitle(ctx_mod), _link2mod_list_column, _link2mod_list_value, string.Empty, string.Empty);
                    report.Messages.Add("Link naar modulesite gemaakt.");
                }
                if (this.LinkExists(ctx_mod, _link2edu_list, link.EduProgramme.GetUrl()))
                {
                    this.LogWarning("Link naar opleidingssite niet gemaakt. Link bestaat al.", report);
                }
                else
                {
                    this.CreateLink(ctx_mod, _link2edu_list, link.EduProgramme.GetUrl(), this.GetTitle(ctx_edu), _link2edu_list_column, _link2edu_list_value, string.Empty, string.Empty);
                    report.Messages.Add("Link naar opleidingssite gemaakt.");
                }
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(link):\n" + e.Message, report); }
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
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.GetUrl());
                this.DeleteLink(ctx_edu, _link2mod_list, link.Module.GetUrl());
                report.Messages.Add("Link naar modulesite verwijderd.");
                ClientContext ctx_mod = this.GetSite(link.Module.GetUrl());
                this.DeleteLink(ctx_mod, _link2edu_list, link.EduProgramme.GetUrl());
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
                report.Messages.Add("Onbepaalde actie op opleiding: id->" + edu.Id + ", code->" + edu.Code + ", naam->" + edu.Name);
                ClientContext ctx = this.GetSite(_edu_url);
                if (this.SiteExists(ctx, edu))
                {
                    report.Messages.Add("De opleidingssite bestaat al.");
                    ClientContext ctx_edu = this.GetSite(edu.GetUrl());
                    string oldname = this.GetTitle(ctx_edu);
                    string newname = edu.GetTitle();
                    if (oldname.Equals(newname))
                    {
                        report.Messages.Add("De naam van de opleiding hoeft niet gewijzigd te worden.");
                        this.LogWarning("Er kon geen actie bepaald worden. Er zijn geen acties genomen.", report);
                    }
                    else
                    {
                        report.Messages.Add("Actie bepaald: De naam van de opleiding moet gewijzigd worden van '" + oldname + "' naar '" + newname + "'.");
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
                ClientContext ctx = this.GetSite(_mod_url);
                if (this.SiteExists(ctx, mod))
                {
                    report.Messages.Add("De modulesite bestaat al.");
                    ClientContext ctx_mod = this.GetSite(mod.GetUrl());
                    string oldname = this.GetTitle(ctx_mod);
                    string newname = mod.GetTitle();
                    if (oldname.Equals(newname))
                    {
                        report.Messages.Add("De naam van de module hoeft niet gewijzigd te worden.");
                        this.LogWarning("Er kon geen actie bepaald worden. Er zijn geen acties genomen.", report);
                    }
                    else
                    {
                        report.Messages.Add("Actie bepaald: De naam van de module moet gewijzigd worden van '" + oldname + "' naar '" + newname + "'.");
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
    }
}
