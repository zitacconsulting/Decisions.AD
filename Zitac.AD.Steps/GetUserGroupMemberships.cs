using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using System;
using System.Collections.Generic;
using System.Collections;
using System.DirectoryServices.Protocols;
using System.Linq;
using System.Security.Principal;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.Text;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Get User Group Memberships", "Integration", "Active Directory", "Zitac", "User")]
    [Writable]
    public class GetUserGroupMembership : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, IDefaultInputMappingStep //, INotifyPropertyChanged
    {

        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool useSSL = true;

        [WritableValue]
        private bool ignoreInvalidCert;

        [WritableValue]
        private bool showOutcomeforNoResults;

        [PropertyClassification(6, "Use Integrated Authentication", new string[] { "Connection" })]
        public bool IntegratedAuthentication
        {
            get { return integratedAuthentication; }
            set
            {
                integratedAuthentication = value;
                //Call OnPropertyChanged method for each property you want to update
                //this.OnPropertyChanged("IntegratedAuthentication");
                //If any of the inputs you want to update are in InputData (not a property),
                //you need to update InputData and shown below.
                this.OnPropertyChanged("InputData");
            }
        }

        [PropertyClassification(7, "Use SSL", new string[] { "Connection" })]
        public bool UseSSL
        {
            get { return useSSL; }
            set
            {
                useSSL = value;
                this.OnPropertyChanged(nameof(UseSSL));
                this.OnPropertyChanged("IgnoreInvalidCert");

            }
        }

        [BooleanPropertyHidden("UseSSL", false)]
        [PropertyClassification(8, "Ignore Certificate Errors", new string[] { "Connection" })]
        public bool IgnoreInvalidCert
        {
            get { return ignoreInvalidCert; }
            set
            {
                ignoreInvalidCert = value;
            }
        }

        [PropertyClassification(1, "Show Outcome for No Results", new string[] { "Outcomes" })]
        public bool ShowOutcomeforNoResults
        {
            get { return showOutcomeforNoResults; }
            set
            {
                showOutcomeforNoResults = value;
                this.OnPropertyChanged("OutcomeScenarios");
            }

        }

        public IInputMapping[] DefaultInputs
        {
            get
            {
                IInputMapping[] inputMappingArray = new IInputMapping[2];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Additional Attributes" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
                return inputMappingArray;
            }
        }

        public DataDescription[] InputData
        {
            get
            {

                List<DataDescription> dataDescriptionList = new List<DataDescription>();
                if (!IntegratedAuthentication)
                {
                    dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(Credentials)), "Credentials"));
                }

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "AD Server"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "User Name"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port", false, true, false));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Additional Attributes", true, true, true));
                return dataDescriptionList.ToArray();
            }
        }


        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(Group), "Found Groups", true)));
                if (ShowOutcomeforNoResults)
                {
                    outcomeScenarioDataList.Add(new OutcomeScenarioData("No Results"));
                }
                outcomeScenarioDataList.Add(new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")));
                return outcomeScenarioDataList.ToArray();
            }
        }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            string UserName = data.Data["User Name"] as string;
            int? Port = (int?)data.Data["Port"];
            List<string> AdditionalAttributes = (data.Data["Additional Attributes"] as string[])?.ToList();

            Credentials ADCredentials = new Credentials();

            if (IntegratedAuthentication)
            {
                ADCredentials.Username = null;
                ADCredentials.Password = null;

            }
            else
            {
                Credentials InputCredentials = data.Data["Credentials"] as Credentials;
                ADCredentials = InputCredentials;
            }

            var Filter = "(&(objectClass=user)(|(sAMAccountName=" + UserName + ")(distinguishedname=" + UserName + ")))";
            try
            {
                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);

                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = LDAPHelper.GetBaseDN(connection);
                List<SearchResultEntry> Results = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "distinguishedname", "primaryGroupID", "objectSid" }).ToList();


                long primaryGroupID = 0;
                string dn = null;
                string sid = null;
                if (Results != null && Results.Count != 0)
                {
                    primaryGroupID = Converters.GetIntProperty(Results[0], "primaryGroupID");
                    dn = Converters.GetStringProperty(Results[0], "distinguishedname");

                    var sidBytes = (byte[])Results[0].Attributes["objectSid"][0];
                    sid = Converters.ConvertSidToString(sidBytes);

                    sid = sid.Remove(sid.LastIndexOf("-", StringComparison.Ordinal) + 1);
                    sid = sid + primaryGroupID.ToString();
                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }


                List<string> BaseAttributeList = Group.GroupAttributes;

                if (AdditionalAttributes != null)
                {
                    BaseAttributeList.AddRange(AdditionalAttributes);
                }
                else
                {
                    AdditionalAttributes = new List<string>();
                }

                List<Group> GroupList = new List<Group>();

                if (primaryGroupID != 0)
                {
                    Filter = "(&(objectClass=group)(objectCategory=group)(ObjectSID=" + sid + "))";
                    Results = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, BaseAttributeList).ToList();
                    if (Results != null && Results.Count != 0)
                    {
                        Group group = new Group(Results[0], AdditionalAttributes);
                        GroupList.Add(group);
                    }

                }

                Filter = "(&(objectClass=group)(objectCategory=group)(member=" + dn + "))";
                Results = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, BaseAttributeList).ToList();

                if (Results != null && Results.Count != 0)
                {
                    foreach (SearchResultEntry Group in Results)
                    {

                        GroupList.Add(new(Group, AdditionalAttributes));
                    }

                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }

                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("Found Groups", (object)GroupList.ToArray());
                return new ResultData("Done", (IDictionary<string, object>)dictionary);

            }
            catch (Exception e)
            {
                string ExceptionMessage = e.ToString();
                return new ResultData("Error", (IDictionary<string, object>)new Dictionary<string, object>()
                {
                {
                    "Error Message",
                    (object) ExceptionMessage
                }
                });

            }
        }
    }
}