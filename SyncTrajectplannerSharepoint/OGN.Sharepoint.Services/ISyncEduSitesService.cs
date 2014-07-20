using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

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
        OperationReport Create([MessageParameter(Name = "Link")]LinkVal link);

        [OperationContract(Name = "WijzigLink")]
        OperationReport Update([MessageParameter(Name = "Wijziging")]UpdateType<LinkVal> change);

        [OperationContract(Name = "VerwijderLink")]
        OperationReport Delete([MessageParameter(Name = "Link")]LinkRef link);

        [OperationContract(Name = "Test")]
        OperationReport Test();

        [OperationContract(Name = "TestFout")]
        OperationReport TestException();
    }

    public interface ICourseTemplate
    {
        string Code {get; set;}
        string Name { get; set; }

        string GetTitle();
        string GetUrl(string baseurl);
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
    public class EduProgramme: ICourseTemplate
    {
        string _code;
        string _name;

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
            return this.Name + " " + this.Code; 
        }

        public string GetUrl(string baseurl)
        {
            return baseurl + this.Code;
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "OpleidingRef")]
    public class EduProgrammeRef: EduProgramme
    {
        public EduProgrammeRef(string code)
            : base()
        {
            base.Code = code;
        }

        [DataMember(Name = "Code", IsRequired=true)]
        new public string Code
        {
            get { return base.Code; }
            set { base.Code = value; }
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "OpleidingVal")]
    public class EduProgrammeVal : EduProgramme
    {

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

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "Module")]
    public class Module: ICourseTemplate
    {
        string _code;
        string _name;

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
            return this.Name + " " + this.Code;
        }

        public string GetUrl(string baseurl)
        {
            return baseurl + this.Code;
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "ModuleRef")]
    public class ModuleRef : Module
    {
        public ModuleRef(string code): base()
        {
            base.Code = code;
        }

        [DataMember(Name = "Code", IsRequired = true)]
        new public string Code
        {
            get { return base.Code; }
            set { base.Code = value; }
        }
    }
    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "ModuleVal")]
    public class ModuleVal : Module
    {

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
        EduProgramme _edu;
        Module _mod;

        public EduProgramme EduProgramme
        {
            get { return _edu; }
            set { _edu = value; }
        }

        public Module Module
        {
            get { return _mod; }
            set { _mod = value; }
        }
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "LinkRef")]
    public class LinkRef : Link
    {
        public LinkRef(EduProgrammeRef edu, ModuleRef mod): base()
        {
            base.EduProgramme = edu;
            base.Module = mod;
        }

        [DataMember(Name = "Opleiding", IsRequired = true)]
        new public EduProgrammeRef EduProgramme
        {
            get { return (EduProgrammeRef)base.EduProgramme; }
            set { base.EduProgramme = value; }
        }

        [DataMember(Name = "Module", IsRequired = true)]
        new public ModuleRef Module
        {
            get { return (ModuleRef)base.Module; }
            set { base.Module = value; }
        }
    }

    [DataContract(Namespace = "http://teamwise.ogn.eu/services/", Name = "LinkVal")]
    public class LinkVal: Link
    {
        [DataMember(Name = "Opleiding", IsRequired = true)]
        new public EduProgrammeVal EduProgramme
        {
            get { return (EduProgrammeVal)base.EduProgramme; }
            set { base.EduProgramme = value; }
        }

        [DataMember(Name = "Module", IsRequired = true)]
        new public ModuleVal Module
        {
            get { return (ModuleVal)base.Module; }
            set { base.Module = value; }
        }
    }
}
