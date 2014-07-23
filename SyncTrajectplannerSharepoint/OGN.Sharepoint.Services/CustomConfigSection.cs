using System;
using System.Configuration;

namespace OGN.Sharepoint.Services
{

    public class SitePermissionsSection : ConfigurationSection
    {
        [ConfigurationProperty("permissions", IsDefaultCollection = false)]
        [ConfigurationCollection(typeof(SitePermissions),
            AddItemName = "add")]
        public SitePermissions Permissions
        {
            get
            {
                SitePermissions perms =
                    (SitePermissions)base["permissions"];
                return perms;
            }
        }

    }

    public class SitePermissions : ConfigurationElementCollection
    {

        protected override ConfigurationElement CreateNewElement()
        {
            return new PermissionBindingConfigElement();
        }

        protected override Object GetElementKey(ConfigurationElement element)
        {
            PermissionBindingConfigElement elem = (PermissionBindingConfigElement)element;
            return elem.SiteGroup+elem.Permission;
        }

        public PermissionBindingConfigElement this[int index]
        {
            get
            {
                return (PermissionBindingConfigElement)BaseGet(index);
            }
            set
            {
                if (BaseGet(index) != null)
                {
                    BaseRemoveAt(index);
                }
                BaseAdd(index, value);
            }
        }
    }

    public class PermissionBindingConfigElement : ConfigurationElement
    {
        [ConfigurationProperty("sitegroup", IsRequired = true)]
        public string SiteGroup
        {
            get
            {
                return (string)this["sitegroup"];
            }
            set
            {
                this["sitegroup"] = value;
            }
        }

        [ConfigurationProperty("permission", IsRequired = true)]
        public string Permission
        {
            get
            {
                return (string)this["permission"];
            }
            set
            {
                this["permission"] = value;
            }
        }

    }
}
