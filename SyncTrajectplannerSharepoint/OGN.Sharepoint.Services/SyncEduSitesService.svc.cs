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

namespace OGN.Sharepoint.Services
{
    public class SyncEduSitesService : ISyncEduSitesService
    {
        private string _home_url = "http://acc-teamwise.ogn.eu/";
        private NetworkCredential _creds;
        private Guid _loi_id = new Guid("6f5a554e-4667-46eb-aabd-9739b1c5a8d6");
        private Guid _cat_id = new Guid("a571a3e6-ac38-442e-8ea5-c98bc3cbaaf2");
        private Guid _mod_id = new Guid("c68d43c8-bee0-4cf4-a1d4-f8329710ae91");
        private Guid _edu_id = new Guid("95adde66-646d-4b27-9242-0972360881bf");
      //  private string _webtemplate = "{581CD673-C0EE-4069-B067-4CB6F61A6BDD}#ModuleSite";
        private string _webtemplate = "{E88B612A-DC40-4ECE-8DE4-964C6F2D7767}#OpleidingModuleSiteTemplate";
        private string _link2edu_list = "Links";
        private string _link2mod_list = "Links";

        //PowerShell: New-EventLog -LogName Application -Source OGN_Sharepoint_Services_SyncEduSitesService
        private string _eventlogsource = "OGN_Sharepoint_Services_SyncEduSitesService";

        public SyncEduSitesService()
        {
            //get settings
                    
            SecureString _pass = new SecureString();
            _pass.AppendChar('V');
            _pass.AppendChar('a');
            _pass.AppendChar('l');
            _pass.AppendChar('u');
            _pass.AppendChar('e');
            _pass.AppendChar('3');
            _pass.AppendChar('l');
            _pass.AppendChar('u');
            _creds = new NetworkCredential("gmichels", _pass, "ad");
        }

        #region ErrorHandlingAndLogging
        private void LogInfo(string msg, OperationReport report)
        {
            EventLog.WriteEntry(_eventlogsource, msg, EventLogEntryType.Information);
            report.Messages.Add(msg);
        }
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
        private void LogWarning(string msg, OperationReport report)
        {
            EventLog.WriteEntry(_eventlogsource, msg, EventLogEntryType.Warning);
            report.Messages.Add(msg);
            report.ResultType = OperationResultType.Warning;
        } 
        #endregion
        
        #region SharepointFunctions
        /*
         * Log warnings using this.LogWarning(msg,report)
         */
        private ClientContext GetSite(string url)
        {
            ClientContext ctx = new ClientContext(url);
            ctx.Credentials = _creds;
            return ctx;
        }

        private void AddTerm(ClientContext ctx, Guid edumod_id, ICourseTemplate edumod)
        {
            TaxonomySession tses = TaxonomySession.GetTaxonomySession(ctx);
            TermStore terms = tses.GetDefaultSiteCollectionTermStore();

            TermGroup loi_group = terms.GetGroup(_loi_id);

            TermSet cat_set = terms.GetTermSet(_cat_id);
            Term edumod_set = cat_set.GetTerm(edumod_id);

            Term term = edumod_set.CreateTerm(edumod.GetTitle(), 1033, Guid.NewGuid());
            ctx.ExecuteQuery();
        }

        private void CreateSite(ClientContext ctx, ICourseTemplate edumod)
        {
            Web site = ctx.Web;

            WebCreationInformation newsite = new WebCreationInformation();
            newsite.WebTemplate = _webtemplate;
            newsite.Title = edumod.GetTitle();
            newsite.Url = edumod.Code;
            newsite.UseSamePermissionsAsParentSite = true;
            site.Webs.Add(newsite);
            ctx.ExecuteQuery();
        }

        private void ChangeTitle(ClientContext ctx, ICourseTemplate edumod)
        {
            Web site = ctx.Web;
            site.Title = edumod.GetTitle();
            site.Update();
            ctx.ExecuteQuery();
        }

        private void CreateLink(ClientContext ctx, string listtitle, ICourseTemplate linkto)
        {
            Web site = ctx.Web;
            List list = site.Lists.GetByTitle(listtitle);
            
            ListItemCreationInformation itemInfo = new ListItemCreationInformation();
            ListItem item = list.AddItem(itemInfo);
            FieldUrlValue url = new FieldUrlValue();
            url.Url = linkto.GetUrl(_home_url);
            url.Description = linkto.GetTitle();
            item["URL"] =  url;
            
            item.Update(); 
            ctx.ExecuteQuery();
        }

        private void UpdateAllLinksToEduOrMod(ClientContext ctx, string listtitle, ICourseTemplate edumod)
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
                string listtitle2 = (listtitle == _link2edu_list) ? _link2mod_list : _link2edu_list;
                this.UpdateLink(ctx2, listtitle2, edumod);
            }
        }

        private void UpdateLink(ClientContext ctx, string listtitle, ICourseTemplate linkto)
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
                if (url.Url == linkto.GetUrl(_home_url))
                {
                    url.Description = linkto.GetTitle();
                    item["URL"] = url;
                    item.Update();
                    break;
                }
            }
            ctx.ExecuteQuery();
        }

        private void DeleteLink(ClientContext ctx, string listtitle, ICourseTemplate linkto)
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
                if (url.Url == linkto.GetUrl(_home_url)) 
                {
                    item.DeleteObject();
                    break;
                }
            } 
            ctx.ExecuteQuery();
        }
        #endregion

        #region ServiceOperations
        /* 
         * SNIPPET FOR OPERATIONS:
            OperationReport report = new OperationReport();
            try
            {
                //operation here
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie ???():\n" + e.Message, report); }
            return report;
         */
        public OperationReport Create(EduProgrammeVal edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak opleiding: code->"+edu.Code+", naam->"+edu.Name);
                ClientContext ctx = this.GetSite(_home_url);
                this.CreateSite(ctx, edu);
                report.Messages.Add("Site gemaakt.");
                this.AddTerm(ctx, _edu_id, edu);
                report.Messages.Add("Term gemaakt.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(edu):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Update(EduProgrammeVal edu)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Wijzig opleidingsnaam: code->" + edu.Code + ", nieuwe naam->" + edu.Name);

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
                report.Messages.Add("Deactiveer opleidingssite: code->" + edu.Code + ", naam->" + edu.Name);
                ClientContext ctx = this.GetSite(edu.GetUrl(_home_url));
                throw new NotImplementedException("Archiveringssysteem nog niet bepaald.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Delete(edu):\n" + e.Message, report); }
            return report;
        }


        public OperationReport Create(ModuleVal mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak module: code->" + mod.Code + ", naam->" + mod.Name);
                ClientContext ctx = this.GetSite(_home_url);
                this.CreateSite(ctx, mod);
                report.Messages.Add("Site gemaakt.");
                this.AddTerm(ctx, _mod_id, mod);
                report.Messages.Add("Term gemaakt.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(mod):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Update(ModuleVal mod)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Wijzig modulenaam: code->" + mod.Code + ", nieuwe naam->" + mod.Name);

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
                report.Messages.Add("Deactiveer modulesite: code->" + mod.Code + ", naam->" + mod.Name);
                ClientContext ctx = this.GetSite(mod.GetUrl(_home_url));
                throw new NotImplementedException("Archiveringssysteem nog niet bepaald.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Delete(mod):\n" + e.Message, report); }
            return report;
        }


        public OperationReport Create(LinkVal link)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Maak links: opl_code->" + link.EduProgramme.Code + ", mod_code->" + link.Module.Code);
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.GetUrl(_home_url));
                this.CreateLink(ctx_edu, _link2mod_list, link.Module);
                report.Messages.Add("Link naar modulesite gemaakt.");
                ClientContext ctx_mod = this.GetSite(link.Module.GetUrl(_home_url));
                this.CreateLink(ctx_mod, _link2edu_list, link.EduProgramme);
                report.Messages.Add("Link naar opleidingssite gemaakt.");
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Create(link):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Update(UpdateType<LinkVal> change)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Vervang link.");
                OperationReport report1 = this.Create(change.To);
                foreach (string msg in report1.Messages) { report.Messages.Add(msg); }
                report.ResultType = report1.ResultType;
                LinkRef linkref = new LinkRef(new EduProgrammeRef(change.From.EduProgramme.Code),new ModuleRef(change.From.Module.Code));
                OperationReport report2 = this.Delete(linkref);
                foreach (string msg in report2.Messages) { report.Messages.Add(msg); }
                report.ResultType = (report2.ResultType==OperationResultType.Warning) ? report2.ResultType : report.ResultType;
            }
            catch (Exception e) { this.LogException(e, "Er is een fout opgetreden tijdens operatie Update(link):\n" + e.Message, report); }
            return report;
        }

        public OperationReport Delete(LinkRef link)
        {
            OperationReport report = new OperationReport();
            try
            {
                report.Messages.Add("Verwijder links: opl_code->" + link.EduProgramme.Code + ", mod_code->" + link.Module.Code);
                ClientContext ctx_edu = this.GetSite(link.EduProgramme.GetUrl(_home_url));
                this.DeleteLink(ctx_edu, _link2mod_list, link.Module);
                report.Messages.Add("Link naar modulesite verwijderd.");
                ClientContext ctx_mod = this.GetSite(link.Module.GetUrl(_home_url));
                this.DeleteLink(ctx_mod, _link2edu_list, link.EduProgramme);
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
    }
}
