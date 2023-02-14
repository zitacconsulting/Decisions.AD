using ActiveDirectory;
using System;
using System.DirectoryServices;

namespace Zitac.AD.Steps
{
    public class PasswordExpiration
    {
        public static Int32 GetADPasswordExpirationPolicy(string ADServer, Credentials ADCredentials)
        {
            try
            {
                string baseLdapPath = string.Format("LDAP://{0}", (object)ADServer);
                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);


                DirectorySearcher searcher = new DirectorySearcher(searchRoot);
                searcher.Filter = "(maxPwdAge=*)";
                searcher.SearchScope = SearchScope.Base;
                searcher.PropertiesToLoad.Add("maxPwdAge");

                SearchResult result = searcher.FindOne();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                searcher.Dispose();

                if (result == null)
                {
                    return 0;
                }

                Int64 val = (Int64)result.Properties["maxPwdAge"][0];
                var maxPwdAge = val / -864000000000; //convert to days
                return (Int32)maxPwdAge;
            }
            catch (Exception e)
            {
                string ExceptionMessage = e.ToString();
                throw new Exception(ExceptionMessage);

            }
        }
    }

}