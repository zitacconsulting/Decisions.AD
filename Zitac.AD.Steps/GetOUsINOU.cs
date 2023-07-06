using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using System;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Linq;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.CoreSteps;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Get OUs In OU", "Integration", "Active Directory", "Zitac", "OU")]
    [Writable]
    public class GetOUsInOU : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, IDefaultInputMappingStep //, INotifyPropertyChanged
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
                IInputMapping[] inputMappingArray = new IInputMapping[3];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Additional Attributes" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
                inputMappingArray[2] = (IInputMapping)new ConstantInputMapping() {InputDataName = "Scope", Value = SearchScope.Subtree};
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port", false, true, false));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "OU DN"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(SearchScope)), "Scope", false, true, true));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Additional Attributes", true, true, true));
                return dataDescriptionList.ToArray();
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(OrganizationalUnit), "OUs", true)));
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
            int? Port = (int?)data.Data["Port"];
            string OUDN = data.Data["OU DN"] as string;
            SearchScope Scope = (SearchScope)data.Data["Scope"];
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

            List<string> BaseAttributeList = OrganizationalUnit.OUAttributes;

            if (AdditionalAttributes != null)
            {
                BaseAttributeList.AddRange(AdditionalAttributes);
            }
            else
            {
                AdditionalAttributes = new List<string>();
            }

            var Filter = "(objectClass=organizationalUnit)";
            try
            {
                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);


                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);

                List<SearchResultEntry> Results = LDAPHelper.GetPagedLDAPResults(connection, OUDN, Scope, Filter, BaseAttributeList).ToList();

                List<OrganizationalUnit> OrganizationalUnits = new List<OrganizationalUnit>();

                if (Results != null && Results.Count != 0)
                {
                    foreach (SearchResultEntry OrganizationalUnit in Results)
                    {

                        OrganizationalUnits.Add(new(OrganizationalUnit, AdditionalAttributes));
                    }

                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }


                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("OUs", (object)OrganizationalUnits.ToArray());
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