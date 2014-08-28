using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Configuration;

namespace OGN.Sharepoint.Services
{
    [ServiceContract(Namespace="http://teamwise.ogn.eu/services/", Name="SyncOpleidingscatalogusService")]
    public interface ISyncEduSitesService
    {
        [OperationContract(Name="MaakOpleiding")]
        OperationReport Create([MessageParameter(Name="Opleiding")]EduProgrammeVal edu);

        [OperationContract(Name="WijzigNaamOpleiding")]
        OperationReport Update([MessageParameter(Name = "WijzigingNaar")]EduProgrammeVal edu);

        [OperationContract(Name="DeactiveerOpleiding")]
        OperationReport Delete([MessageParameter(Name = "Opleiding")]EduProgrammeRef edu);

        [OperationContract(Name = "MaakModule")]
        OperationReport Create([MessageParameter(Name = "Module")]ModuleVal mod);

        [OperationContract(Name = "WijzigNaamModule")]
        OperationReport Update([MessageParameter(Name = "WijzigingNaar")]ModuleVal mod);

        [OperationContract(Name = "DeactiveerModule")]
        OperationReport Delete([MessageParameter(Name = "Module")]ModuleRef mod);

        [OperationContract(Name = "MaakLink")]
        OperationReport Create([MessageParameter(Name = "Link")]Link link);

        [OperationContract(Name = "WijzigLink")]
        OperationReport Update([MessageParameter(Name = "Wijziging")]UpdateType<Link> change);

        [OperationContract(Name = "VerwijderLink")]
        OperationReport Delete([MessageParameter(Name = "Link")]Link link);

        [OperationContract(Name = "Test")]
        OperationReport Test();

        [OperationContract(Name = "TestFout")]
        OperationReport TestException();

        [OperationContract(Name = "DoeOnbepaaldeActieOpleiding")]
        OperationReport DoUndeterminedAction([MessageParameter(Name = "Opleiding")]EduProgrammeVal edu);

        [OperationContract(Name = "DoeOnbepaaldeActieModule")]
        OperationReport DoUndeterminedAction([MessageParameter(Name = "Module")]ModuleVal mod);
    }

    public interface IEduModSite
    {
        /// <summary>
        /// technical id
        /// </summary>
        string Id { get; set; }
        /// <summary>
        /// business id
        /// </summary>
        string Code {get; set;}
        /// <summary>
        /// name of programme or module
        /// </summary>
        string Name { get; set; }
        
        /// <summary>
        /// the title of the site
        /// </summary>
        /// <returns></returns>
        string GetTitle();
        /// <summary>
        /// the relative url of the site
        /// </summary>
        /// <returns></returns>
        string GetSiteName(); 
        /// <summary>
        /// the full url of the site
        /// </summary>
        /// <param name="baseurl">the url of the parent site</param>
        /// <returns></returns>
        string GetUrl(); 
    }

    [DataContract(Namespace="http://teamwise.ogn.eu/services/", Name="Wijziging {0}")]
    public class UpdateType<T>
    {
        T _from;
        T _to;

        [DataMember(Name="Van")]
        public T From
        {
            get { return _from; }
            set { _from = value; }
        }

        [DataMember(Name="Naar")]
        public T To
        {
            get { return _to; }
            set { _to = value; }
        }

    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "ActieResultaat")]
    public enum OperationResultType
    {
        //
        // Summary:
        //     This indicates a successful operation with warnings.
        [EnumMember(Value="Waarschuwing")]
        Warning = 2,
        //
        // Summary:
        //     This indicates a successful operation.
        [EnumMember]
        OK = 1,
    }

    [CollectionDataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "Trace", ItemName = "Bericht")]
    public class Messages : List<string> 
    {
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "ActieRapport")]
    public class OperationReport
    {
        
        OperationResultType _type = OperationResultType.OK;
        Messages _msgs = new Messages();
        
        [DataMember(Name = "ActieResultaat")]
        public OperationResultType ResultType
        {
            get { return _type; }
            set { _type = value; }
        }

        [DataMember(Name = "Berichten")]
        public Messages Messages
        {
            get { return _msgs; }
            set { _msgs = value; }
        }
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name="Opleiding")]
    public class EduProgramme: IEduModSite
    {
        string _id;
        string _code;
        string _name;

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public string Code
        {
            get { return _code; }
            set { _code = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string GetTitle()
        {
            return (string.Empty.Equals(this.Code))?this.Name:this.Name + " " + this.Code; 
        }

        public string GetUrl()
        {
            return ConfigurationManager.AppSettings["sp.sitecollection:edu:url"] + this.GetSiteName();
        }


        public string GetSiteName()
        {
            return System.Web.HttpUtility.UrlEncode(this.Id);
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "OpleidingRef")]
    public class EduProgrammeRef: EduProgramme
    {
        public EduProgrammeRef(string id)
            : base()
        {
            base.Id = id;
        }

        [DataMember(Name = "Id", IsRequired = true)]
        new public string Id
        {
            get { return base.Id; }
            set { base.Id = value; }
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "OpleidingVal")]
    public class EduProgrammeVal : EduProgramme
    {
        [DataMember(Name = "Id", IsRequired = true)]
        new public string Id
        {
            get { return base.Id; }
            set { base.Id = value; }
        }

        [DataMember(Name = "Code")]
        new public string Code
        {
            get { return base.Code; }
            set { base.Code = value; }
        }

        [DataMember(Name = "Naam", IsRequired = true)]
        new public string Name
        {
            get { return base.Name; }
            set { base.Name = value; }
        }
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "Module")]
    public class Module: IEduModSite
    {
        string _id;
        string _code;
        string _name;

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public string Code
        {
            get { return _code; }
            set { _code = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string GetTitle()
        {
            return (string.Empty.Equals(this.Code)) ? this.Name : this.Name + " " + this.Code; 
        }

        public string GetUrl()
        {
            return ConfigurationManager.AppSettings["sp.sitecollection:mod:url"] + this.GetSiteName();
        }

        public string GetSiteName()
        {
            return System.Web.HttpUtility.UrlEncode(this.Id);
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "ModuleRef")]
    public class ModuleRef : Module
    {
        public ModuleRef(string id): base()
        {
            base.Id = id;
        }

        [DataMember(Name = "Id", IsRequired = true)]
        new public string Id
        {
            get { return base.Id; }
            set { base.Id = value; }
        }

    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "ModuleVal")]
    public class ModuleVal : Module
    {
        [DataMember(Name = "Id", IsRequired = true)]
        new public string Id
        {
            get { return base.Id; }
            set { base.Id = value; }
        }

        [DataMember(Name = "Code", IsRequired = true)]
        new public string Code
        {
            get { return base.Code; }
            set { base.Code = value; }
        }

        [DataMember(Name = "Naam", IsRequired = true)]
        new public string Name
        {
            get { return base.Name; }
            set { base.Name = value; }
        }
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "Link")]
    public class Link
    {
        EduProgrammeRef _edu;
        ModuleRef _mod;

        [DataMember(Name = "Opleiding", IsRequired = true)]
        public EduProgrammeRef EduProgramme
        {
            get { return _edu; }
            set { _edu = value; }
        }

        [DataMember(Name = "Module", IsRequired = true)]
        public ModuleRef Module
        {
            get { return _mod; }
            set { _mod = value; }
        }
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "SiteIdPaar")]
    public class PairOfSiteId
    {
        string _id1;
        string _id2;

        [DataMember(Name = "SiteId1", IsRequired = true)]
        public string SiteId1
        {
            get { return _id1; }
            set { _id1 = value; }
        }

        [DataMember(Name = "SiteId2", IsRequired = true)]
        public string SiteId2
        {
            get { return _id2; }
            set { _id2 = value; }
        }
    }
}
