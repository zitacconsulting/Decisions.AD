using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using System;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Linq;
using System.Security.Principal;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.CoreSteps;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Get Computers In Group", "Integration", "Active Directory", "Zitac", "Group")]
    [Writable]
    public class GetComputersInGroup : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, IDefaultInputMappingStep //, INotifyPropertyChanged
    {

        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool useSSL = true;

        [WritableValue]
        private bool ignoreInvalidCert;

        [WritableValue]
        private bool showOutcomeforNoResults;

        [WritableValue]
        private bool recursiveSearch;

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

        [PropertyClassification(new string[] { "Recursive Search" })]
        public bool RecursiveSearch
        {
            get { return recursiveSearch; }
            set
            {
                recursiveSearch = value;
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port",false, true, false));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Group name or DN"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Additional Attributes", true, true, true));
                return dataDescriptionList.ToArray();
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(Computer), "Computers", true)));
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
            string GroupName = data.Data["Group name or DN"] as string;
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

            List<string> BaseAttributeList = Computer.ComputerAttributes;

            if (AdditionalAttributes != null)
            {
                BaseAttributeList.AddRange(AdditionalAttributes);
            }
            else
            {
                AdditionalAttributes = new List<string>();
            }

            var Filter = "(&(objectClass=group)(objectCategory=group)(|(sAMAccountName=" + GroupName + ")(distinguishedname=" + GroupName + ")))";

            try
            {
                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);

                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = LDAPHelper.GetBaseDN(connection);
                List<SearchResultEntry> Group = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "objectSid", "distinguishedname"}).ToList();

                SearchResultEntry FoundGroup = null;

                if (Group != null && Group.Count != 0)
                {
                    FoundGroup = Group[0];
                }
                else
                {
                    if (ShowOutcomeforNoResults)
                    {
                        return new ResultData("No Results");
                    }
                    throw new Exception(string.Format("Unable to find group with name: '{0}' in the AD", (object)GroupName));
                }

                List<Computer> ComputerList = new List<Computer>();

                byte[] sidBytes = (byte[])FoundGroup.Attributes["objectSid"][0];
                string SID = Converters.ConvertSidToString(sidBytes);
                string RID = SID.Substring(SID.LastIndexOf("-", StringComparison.Ordinal) + 1);

                if (RecursiveSearch)
                {
                    Filter = "(&(objectClass=computer)(|(memberOf:1.2.840.113556.1.4.1941:=" + FoundGroup.Attributes["distinguishedname"][0] + ")(primaryGroupId=" + RID + ")))";
                }
                else
                {
                    Filter = "(&(objectClass=computer)(|(memberOf=" + FoundGroup.Attributes["distinguishedname"][0] + ")(primaryGroupId=" + RID + ")))";
                }

                var GroupMembers = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, BaseAttributeList);
                int MaxPasswordAge = LDAPHelper.GetADPasswordExpirationPolicy(connection, BaseDN);
                foreach (SearchResultEntry Current in GroupMembers)
                {
                    ComputerList.Add(new Computer(Current, AdditionalAttributes));
                }


                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("Computers", (object)ComputerList.ToArray());
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