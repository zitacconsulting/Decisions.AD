using ActiveDirectory;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Service.Debugging.DebugData;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security.Permissions;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using DecisionsFramework.ServiceLayer;
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

        [PropertyClassification(8, "Use Integrated Authentication", new string[] {"Integrated Authentication"})]
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

 
            public DataDescription[] InputData
            {
                    get {
                        
                        List<DataDescription> dataDescriptionList = new List<DataDescription>();
                            if(!IntegratedAuthentication)
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (Credentials)), "Credentials"));
                            }
                            
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "AD Server"));

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
           
            Credentials ADCredentials = new Credentials();

            if(IntegratedAuthentication)
            {
                ADCredentials.ADUsername = null;
                ADCredentials.ADPassword = null;
                  
            }
            else
            {
                Credentials InputCredentials = data.Data["Credentials"] as Credentials;
                ADCredentials = InputCredentials;
            }

            try
            {
                string baseLdapPath = string.Empty;
                baseLdapPath = string.Format("LDAP://{0}", (object) ADServer);

                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);
                DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
                directorySearcher.Filter = "(&(objectCategory=computer)(objectClass=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))";

                SearchResultCollection All = directorySearcher.FindAll();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                directorySearcher.Dispose();

                List<DomainController> Results = new List<DomainController>();
                if (All != null && All.Count != 0)
                {
                    foreach (SearchResult Current in All)
                    {
                        Results.Add(new DomainController(Current));
                    }
                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }
                

                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("DomainControllers", (object) Results.ToArray());
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