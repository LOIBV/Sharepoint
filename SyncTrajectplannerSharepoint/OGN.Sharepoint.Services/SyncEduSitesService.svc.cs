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

namespace OGN.Sharepoint.Services
{
    public class SyncEduSitesService : ISyncEduSitesService
    {
        //web.config settings (see constructor)
        private string _home_url;
        private NetworkCredential _creds;
        private Guid _loi_id;
        private Guid _cat_id;
        private Guid _mod_id;
        private Guid _edu_id;
        private string _modtemplate;
        private string _edutemplate;
        private string _link2edu_list;
        private string _link2mod_list;

        //PowerShell: New-EventLog -LogName Application -Source OGN_Sharepoint_Services_SyncEduSitesService
        private string _eventlogsource = "OGN_Sharepoint_Services_SyncEduSitesService";

        public SyncEduSitesService()
        {
            //get web.config settings       
            SecureString _pass = new SecureString();
            _pass.AppendChar('V');
            _pass.AppendChar('a');
            _pass.AppendChar('l');
            _pass.AppendChar('u');
            _pass.AppendChar('e');
            _pass.AppendChar('3');
            _pass.AppendChar('l');
            _pass.AppendChar('u');
            //the credentials of the application pool are used.
            _creds = CredentialCache.DefaultNetworkCredentials; 
            //_creds = new NetworkCredential("gmichels", _pass, "ad"); //for testing

            _home_url = ConfigurationManager.AppSettings["sp.sitecollection:url"];
            _loi_id = new Guid(ConfigurationManager.AppSettings["sp.termstore:id"]);
            _cat_id = new Guid(ConfigurationManager.AppSettings["sp.termstore.termset:id"]);
            _mod_id = new Guid(ConfigurationManager.AppSettings["sp.termstore.termset.modset:id"]);
            _edu_id = new Guid(ConfigurationManager.AppSettings["sp.termstore.termset.eduset:id"]);
            _modtemplate = ConfigurationManager.AppSettings["sp.modsite:template"];
            _link2edu_list = ConfigurationManager.AppSettings["sp.modsite:list2edu"];
            _edutemplate = ConfigurationManager.AppSettings["sp.edusite:template"];
            _link2mod_list = ConfigurationManager.AppSettings["sp.edusite:list2mod"];
            _eventlogsource = ConfigurationManager.AppSettings["eventlogsource"];
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
            EventLog.WriteEntry(_eventlogsource, msg+"\n\n"+tracedump, EventLogEntryType.Warning);
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
            Web site = ctx.Web;

            WebCreationInformation newsite = new WebCreationInformation();
            newsite.WebTemplate = template;
            newsite.Title = edumod.GetTitle();
            newsite.Url = edumod.GetSiteName();
            newsite.UseSamePermissionsAsParentSite = true;
            site.Webs.Add(newsite);
            ctx.ExecuteQuery();
        }

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

            return 0<site.Webs.Count(subsite => subsite.Url.EndsWith("/" + edumod.GetSiteName()));
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
        private void ChangePermissions(ClientContext ctx)
        {
            Web site = ctx.Web;
            site.BreakRoleInheritance(false,false);
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
            
            SitePermissionsSection config = (SitePermissionsSection)ConfigurationManager.GetSection("sp.sitepermissions.deactivate");

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
        private void CreateLink(ClientContext ctx, string listtitle, string linktourl, string linktodescr)
        {
            Web site = ctx.Web;
            List list = site.Lists.GetByTitle(listtitle);
            
            ListItemCreationInformation itemInfo = new ListItemCreationInformation();
            ListItem item = list.AddItem(itemInfo);
            FieldUrlValue url = new FieldUrlValue();
            url.Url = linktourl; //linkto.GetUrl(_home_url);
            url.Description = linktodescr; //linkto.GetTitle();
            item["URL"] =  url;
            
            item.Update(); 
            ctx.ExecuteQuery();
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
                if (url.Url.Equals(linkto.GetUrl(_home_url)))
                {
                    url.Description = linkto.GetTitle();
                    item["URL"] = url;
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
                FieldUrlValue url = (FieldUrlValue) item["URL"];
                if (url.Url.Equals(linktourl)) 
                {
                    item.DeleteObject();
                    break;
                }
            } 
            ctx.ExecuteQuery();
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
                report.Messages.Add("Maak opleiding: id->"+edu.Id+", code->"+edu.Code+", naam->"+edu.Name);
                ClientContext ctx = this.GetSite(_home_url);
                if (this.SiteExists(ctx, edu))
                {
                    this.LogWarning("Site niet gemaakt. Opleiding bestaat al.", report);
                }
                else
                {
                    this.CreateSite(ctx, edu, _edutemplate);
                    report.Messages.Add("Site gemaakt.");
                    this.AddTerm(ctx, _edu_id, edu);
                    report.Messages.Add("Term gemaakt.");
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

                ClientContext ctx = this.GetSite(edu.GetUrl(_home_url));
                this.ChangeTitle(ctx, edu);
                report.Messages.Add("Site titel gewijzigd.");
                this.AddTerm(ctx, _edu_id, edu);
                report.Messages.Add("Term gemaakt.");
                this.UpdateAllLinksToEduOrMod(ctx, _link2mod_list, edu);
                report.Messages.Add("Beschrijvingen van links naar deze site gewijzigd.");
                
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
                ClientContext ctx = this.GetSite(edu.GetUrl(_home_url));
                this.ChangePermissions(ctx);
                report.Messages.Add("Permissies ingetrokken.");
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
                ClientContext ctx = this.GetSite(_home_url);
                if (this.SiteExists(ctx, mod))
                {
                    this.LogWarning("Site niet gemaakt. Module bestaat al.", report);
                }
                else
                {
                    this.CreateSite(ctx, mod, _modtemplate);
                    report.Messages.Add("Site gemaakt.");
                    this.AddTerm(ctx, _mod_id, mod);
                    report.Messages.Add("Term gemaakt.");
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

                ClientContext ctx = this.GetSite(mod.GetUrl(_home_url));
                this.ChangeTitle(ctx, mod);
                report.Messages.Add("Site titel gewijzigd.");
                this.AddTerm(ctx, _mod_id, mod);
                report.Messages.Add("Term gemaakt.");
                this.UpdateAllLinksToEduOrMod(ctx, _link2edu_list, mod);
                report.Messages.Add("Beschrijvingen van links naar deze site gewijzigd.");
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
                ClientContext ctx = this.GetSite(mod.GetUrl(_home_url));
                this.ChangePermissions(ctx);
                report.Messages.Add("Permissies ingetrokken.");
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
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.GetUrl(_home_url));
                ClientContext ctx_mod = this.GetSite(link.Module.GetUrl(_home_url));

                if (this.LinkExists(ctx_edu, _link2mod_list, link.Module.GetUrl(_home_url)))
                {
                    this.LogWarning("Link naar modulesite niet gemaakt. Link bestaat al.", report);
                }
                else
                {
                    this.CreateLink(ctx_edu, _link2mod_list, link.Module.GetUrl(_home_url), this.GetTitle(ctx_mod));
                    report.Messages.Add("Link naar modulesite gemaakt.");
                }
                if (this.LinkExists(ctx_mod, _link2edu_list, link.EduProgramme.GetUrl(_home_url)))
                {
                    this.LogWarning("Link naar opleidingssite niet gemaakt. Link bestaat al.", report);
                }
                else
                {
                    this.CreateLink(ctx_mod, _link2edu_list, link.EduProgramme.GetUrl(_home_url), this.GetTitle(ctx_edu));
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
                report.ResultType = (report2.ResultType==OperationResultType.Warning) ? report2.ResultType : report.ResultType;
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
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.GetUrl(_home_url));
                this.DeleteLink(ctx_edu, _link2mod_list, link.Module.GetUrl(_home_url));
                report.Messages.Add("Link naar modulesite verwijderd.");
                ClientContext ctx_mod = this.GetSite(link.Module.GetUrl(_home_url));
                this.DeleteLink(ctx_mod, _link2edu_list, link.EduProgramme.GetUrl(_home_url));
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
                report.Messages.Add("Dit is de Test operatie van deze service.");
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
                ClientContext ctx = this.GetSite(_home_url);
                if (this.SiteExists(ctx, edu))
                {
                    report.Messages.Add("De opleidingssite bestaat al.");
                    ClientContext ctx_edu = this.GetSite(edu.GetUrl(_home_url));
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
                ClientContext ctx = this.GetSite(_home_url);
                if (this.SiteExists(ctx, mod))
                {
                    report.Messages.Add("De modulesite bestaat al.");
                    ClientContext ctx_mod = this.GetSite(mod.GetUrl(_home_url));
                    string oldname = this.GetTitle(ctx_mod);
                    string newname = mod.GetTitle();
                    if (oldname.Equals(newname))
                    {
                        report.Messages.Add("De naam van de module hoeft niet gewijzigd te worden.");
                        this.LogWarning("Er kon geen actie bepaald worden. Er zijn geen acties genomen.", report);
                    }
                    else
                    {
                        report.Messages.Add("Actie bepaald: De naam van de module moet gewijzigd worden van '"+oldname+"' naar '"+newname+"'.");
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
