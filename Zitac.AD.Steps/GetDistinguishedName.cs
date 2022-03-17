using ActiveDirectory;
using System.ComponentModel;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;



namespace Zitac.AD.Steps;

public class GetDistinguishedName
{

        public static string GetObjectDistinguishedName(
        ADObjectType objectCls,
        string objectName,
        [PropertyClassification(0, "System User Name", new string[] {"Inputs", "LDAP Server Settings"})] string systemUserName,
        [PasswordText, PropertyClassification(1, "System Password", new string[] {"Inputs", "LDAP Server Settings"})] string systemPassword,
        [PropertyClassification(2, "LDAP Server Address", new string[] {"Inputs", "LDAP Server Settings"})] string ldapServerAddress)
        {
            string str = string.Empty;
            string baseLdapPath = ConvertLdapServerAddressToBaseLdapPath(ldapServerAddress);
            DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, systemUserName, systemPassword);
            DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
            switch (objectCls)
            {
                case ADObjectType.User:
                directorySearcher.Filter = "(&(objectClass=user)((sAMAccountName=" + objectName + ")))";
                break;
                case ADObjectType.Group:
                directorySearcher.Filter = "(&(objectClass=group)(|(cn=" + objectName + ")(dn=" + objectName + ")))";
                break;
                case ADObjectType.Computer:
                directorySearcher.Filter = "(&(objectClass=computer)(|(cn=" + objectName + ")(dn=" + objectName + ")))";
                break;
            }
            SearchResult one = directorySearcher.FindOne();
            if (one != null)
            {
                DirectoryEntry directoryEntry = one.GetDirectoryEntry();
                str = string.Format("{0}/{1}", (object) baseLdapPath, directoryEntry.Properties["distinguishedName"].Value);
            }
            if (searchRoot != null)
            {
                searchRoot.Close();
                searchRoot.Dispose();
            }
            directorySearcher.Dispose();
            return str;
        }

    private static string ConvertLdapServerAddressToBaseLdapPath(string ldapServerAddress) => !ldapServerAddress.StartsWith("LDAP://", StringComparison.InvariantCultureIgnoreCase) ? string.Format("LDAP://{0}", (object) ldapServerAddress) : ldapServerAddress;

    public enum ADObjectType
    {
      User,
      Group,
      Computer,
      OrganizationalUnit,
    }
}
