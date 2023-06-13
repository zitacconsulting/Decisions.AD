using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using System;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Linq;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.ComponentModel;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Get Domain Controllers", "Integration", "Active Directory", "Zitac")]
    [Writable]
    public class GetDomainControllers : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer , INotifyPropertyChanged
    {
     
        [WritableValue]
        private bool showOutcomeforNoResults;

        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool useSSL = true;

        [WritableValue]
        private bool ignoreInvalidCert;

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

        [PropertyClassification(1, "Show Outcome for No Results", new string[] {"Outcomes"})]
        public bool ShowOutcomeforNoResults
        {
            get {return showOutcomeforNoResults; }
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
                IInputMapping[] inputMappingArray = new IInputMapping[1];
                inputMappingArray[3] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
                return inputMappingArray;
            }
        }
            public DataDescription[] InputData
            {
                    get {
                        
                        List<DataDescription> dataDescriptionList = new List<DataDescription>();
                            if(!IntegratedAuthentication)
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (Credentials)), "Credentials"));
                            }
                            
                            dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "AD Server"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port",false, true, false));

                            return dataDescriptionList.ToArray();                                              
                        }
            }
  
            public override OutcomeScenarioData[] OutcomeScenarios {
                get {
                    List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();
                    
                    outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(DomainController), "DomainControllers",true)));
                    if (ShowOutcomeforNoResults) {
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
           
            Credentials ADCredentials = new Credentials();

            if(IntegratedAuthentication)
            {
                ADCredentials.Username = null;
                ADCredentials.Password = null;
                  
            }
            else
            {
                Credentials InputCredentials = data.Data["Credentials"] as Credentials;
                ADCredentials = InputCredentials;
            }
            List<string> BaseAttributeList = DomainController.DCAttributes;

            var Filter = "(&(objectCategory=computer)(objectClass=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))";
            try
            {

                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);


                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);

                List<SearchResultEntry> Results = LDAPHelper.GetPagedLDAPResults(connection, null, SearchScope.Subtree, Filter, BaseAttributeList).ToList();

                List<DomainController> DomainControllers = new List<DomainController>();

                if (Results != null && Results.Count != 0)
                {
                    foreach (SearchResultEntry DomainController in Results)
                    {

                        DomainControllers.Add(new(DomainController));
                    }

                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }


                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("DomainControllers", (object) DomainControllers.ToArray());
                return new ResultData("Done", (IDictionary<string, object>) dictionary);



            }
            catch (Exception e)
            {
                string ExceptionMessage = e.ToString();
                return new ResultData("Error", (IDictionary<string, object>) new Dictionary<string, object>()
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