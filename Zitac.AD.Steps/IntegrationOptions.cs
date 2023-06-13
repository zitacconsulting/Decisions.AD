using System;

namespace Zitac.AD.Steps
{
    public class IntegrationOptions
    {
        public readonly string Host;
        public readonly int Port;
        public readonly Credentials Credentials;
        public readonly bool UseSSL;
        public readonly bool IgnoreInvalidCert;
        public readonly bool IntegratedAuthentication;

        public IntegrationOptions(string host, int? port, Credentials credentials, bool useSSL, bool ignoreInvalidCert, bool integratedAuthentication)
        {
            Host = host ?? throw new ArgumentNullException(nameof(host));
            if(port == 0 || port == null) {
                if(useSSL) {
                    Port = 636;
                }
                else {
                    Port = 389;
                }
            }
            else {
                Port = (int)port;
            }
            Credentials = credentials;
            UseSSL = useSSL;
            IgnoreInvalidCert = ignoreInvalidCert;
            IntegratedAuthentication = integratedAuthentication;
        }
    }
        public class QueryOptions
    {
        public string SearchBase;
        public string Filter;
        public readonly int ResultPageSize;
        public readonly string[] TargetAttributes;

        public QueryOptions(string searchBase, string filter, int resultPageSize, string[] targetAttribute)
        {
            SearchBase = searchBase ?? throw new ArgumentNullException(nameof(searchBase));
            Filter = filter ?? throw new ArgumentNullException(nameof(filter));
            ResultPageSize = resultPageSize;
            TargetAttributes = targetAttribute;
        }
    }
}