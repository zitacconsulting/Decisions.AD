using System.DirectoryServices.Protocols;
using System;
using System.Net;
using System.Collections.Generic;
using DecisionsFramework.Design.Flow;

namespace Zitac.AD.Steps
{

    public class LDAPHelper
    {
        public static LdapConnection GenerateLDAPConnection(IntegrationOptions Options)
        {
            var di = new LdapDirectoryIdentifier(server: Options.Host, Options.Port, fullyQualifiedDnsHostName: true, connectionless: false);
            var connection = new LdapConnection(di);
            if (Options.IntegratedAuthentication)
            {
                connection.AuthType = AuthType.Negotiate;
            }
            else
            {
                connection.Credential = new NetworkCredential(Options.Credentials.Username, Options.Credentials.Password, Options.Credentials.Domain);
                connection.AuthType = AuthType.Basic;

            }
            connection.SessionOptions.ReferralChasing = ReferralChasingOptions.None;
            connection.SessionOptions.ProtocolVersion = 3; //Setting LDAP Protocol to latest version
            connection.Timeout = TimeSpan.FromMinutes(1);
            connection.AutoBind = true;
            if (Options.UseSSL)
            {
                connection.SessionOptions.SecureSocketLayer = Options.UseSSL;

                if (Options.IgnoreInvalidCert)
                {
                    if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                    {
                        connection.SessionOptions.VerifyServerCertificate = (ldapConnection, certificate) => true;
                    }
                    else if (!Environment.GetEnvironmentVariables().Contains("LDAPTLS_REQCERT"))
                    {
                        throw new Exception("LDAPS certificate validation is disabled in config, but LDAPTLS_REQCERT environmental variable is not set. On non-Windows environments certificate validation must be disabled by setting environmental variable LDAPTLS_REQCERT to 'never'");
                    }
                }
            }


            connection.Bind();
            return connection;
        }
        public static SearchResultEntry[] GetPagedLDAPResults(LdapConnection connection, string BaseDN,SearchScope ScopeToSearch, string Filter, List<string> AttributeList)
        {
            var pageSize = 1000;
            var cookie = new byte[0];
            var results = new List<SearchResultEntry>();

            if (string.IsNullOrEmpty(BaseDN))
            {
                BaseDN = GetBaseDN(connection);
            }
            do
            {
                var searchRequest = new SearchRequest(
                    BaseDN,
                    Filter,
                    ScopeToSearch,
                    AttributeList.ToArray()
                );

                var pageRequestControl = new PageResultRequestControl(pageSize);

                pageRequestControl.Cookie = cookie;
                searchRequest.Controls.Add(pageRequestControl);

                var searchResponse = (SearchResponse)connection.SendRequest(searchRequest);

                foreach (SearchResultEntry entry in searchResponse.Entries)
                {
                    results.Add(entry);
                }

                var pageResponseControl = (PageResultResponseControl)searchResponse.Controls[0];
                cookie = pageResponseControl.Cookie;

            } while (cookie != null && cookie.Length > 0);
            return results.ToArray();
        }

        public static string GetBaseDN(LdapConnection connection)
        {

            var searchRequest = new SearchRequest(
                "",
                "(objectClass=*)",
                SearchScope.Base,
                new string[] { "defaultNamingContext" }
            );

            var searchResponse = (SearchResponse)connection.SendRequest(searchRequest);
            SearchResultEntry searchResultEntry = searchResponse.Entries[0];

            return searchResultEntry.Attributes["defaultNamingContext"][0].ToString();
        }

        public static int GetADPasswordExpirationPolicy(LdapConnection connection, string BaseDN)
        {
            if (string.IsNullOrEmpty(BaseDN))
            {
                BaseDN = GetBaseDN(connection);
            }
            var searchRequest = new SearchRequest(
                BaseDN,
                "(objectClass=domain)",
                SearchScope.Base,
                new string[] { "maxPWDAge" }
            );

            var searchResponse = (SearchResponse)connection.SendRequest(searchRequest);
            SearchResultEntry searchResultEntry = searchResponse.Entries[0];
            long raw = Convert.ToInt64(searchResultEntry.Attributes["maxPwdAge"][0]);
            var maxPwdAge = raw / -864000000000; //convert to days

            return (Int32)maxPwdAge;
        }

        public static DirectoryAttributeModification CreateAttributeModification(StepStartData data, AttributeValues Attribute) {
                
                if (data.Data[Attribute.Parameter] != null && data.Data[Attribute.Parameter].GetType().ToString() == "System.DateTime")
                {
                    DirectoryAttributeModification AttributeModification =  new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = Attribute.Attribute };
                    AttributeModification.Add(Convert.ToString(((DateTime)data.Data["Account Expires"]).ToFileTimeUtc()));
                    
                    return AttributeModification;
                }
                else if (data.Data[Attribute.Parameter] != null && (data.Data[Attribute.Parameter]).ToString().Length != 0)
                {
                    DirectoryAttributeModification AttributeModification =  new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = Attribute.Attribute };
                    AttributeModification.Add((string)data.Data[Attribute.Parameter]);
                    return AttributeModification;
                }
                else if (data.Data.ContainsKey(Attribute.Parameter) && (data.Data[Attribute.Parameter]) == null)
                {
                    DirectoryAttributeModification AttributeModification =  new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = Attribute.Attribute };
                    return AttributeModification;
                }
                else 
                {
                    return null;
                }
        }
    }

}