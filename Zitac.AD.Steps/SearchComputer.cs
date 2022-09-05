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
    [AutoRegisterStep("Search Computer", "Integration", "Active Directory", "Zitac", "Computer")]
    [Writable]
    public class SearchComputer : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer , INotifyPropertyChanged,  IDefaultInputMappingStep
    {
     
        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool combineFiltersUsingAnd;

        [WritableValue]
        private bool showOutcomeforNoResults;

        [WritableValue]
        private SearchParameters[] qParams;

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

        [PropertyClassification(9, "Combine Filters Using And", new string[] {"Search Definition"})]
        public bool CombineFiltersUsingAnd
        {
            get {return combineFiltersUsingAnd; }
            set {combineFiltersUsingAnd = value;}

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

        [PropertyClassification(10, "Search Criteria", new string[] {"Search Definition"})]
        public SearchParameters[] QueryParams
            {
                get
                {
                    return qParams;
                }
                set
                {
                    qParams = value;
                    this.OnPropertyChanged("InputData");
                    this.OnPropertyChanged(nameof (QueryParams));
                }
            }


            public IInputMapping[] DefaultInputs
            {
            get
            {
                IInputMapping[] inputMappingArray = new IInputMapping[2];
                inputMappingArray[0] = (IInputMapping) new IgnoreInputMapping() { InputDataName = "Search Base (DN)" };
                inputMappingArray[1] = (IInputMapping) new IgnoreInputMapping() { InputDataName = "Additional Attributes" };
                return inputMappingArray;
            }
            }

            protected SearchParameters[] GetSearchParameters()
            {
            if (this.qParams == null || this.qParams.Length == 0)
                return this.qParams;
            List<SearchParameters> queryParametersList = new List<SearchParameters>();
            foreach (SearchParameters qParam in this.qParams)
            {
                if (!string.IsNullOrEmpty(qParam.FieldName))
                queryParametersList.Add(qParam);
            }
            return queryParametersList.ToArray();
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
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Search Base (DN)"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Additional Attributes", true, true, true));

                            SearchParameters[] ParametersList = this.GetSearchParameters();
                            if (ParametersList != null && ParametersList.Length != 0)
                            {
                                foreach (SearchParameters CurrParameter in ParametersList)
                                {
                                    dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (object)), CurrParameter.Alias));
                                }
                            }

                            return dataDescriptionList.ToArray();                                              
                        }
            }
  
            public override OutcomeScenarioData[] OutcomeScenarios {
                get {
                    List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();
                    
                    outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(Computer), "FoundComputers",true)));
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
            string BaseSearch = data.Data["Search Base (DN)"] as string;
            string[] AdditionalAttributes = data.Data["Additional Attributes"] as string[];

            string Filter = string.Empty;

            SearchParameters[] ParametersList = this.GetSearchParameters();
                if (ParametersList != null && ParametersList.Length != 0)
                {
                    if (CombineFiltersUsingAnd)
                        Filter = "(&";
                    else
                        Filter = "(|";
                    foreach (SearchParameters CurrParameter in ParametersList)
                    {
                        object SearchValue = data.Data[CurrParameter.Alias] as object;
                        switch (CurrParameter.MatchCriteria)
                         {
                             case "Equals":
                             Filter += "(" + CurrParameter.FieldName + "=" + SearchValue + ")";
                             break;

                             case "Contains":
                             Filter += "(" + CurrParameter.FieldName + "=*" + SearchValue + "*)";
                             break;

                             case "DoesNotContain":
                             Filter += "(!(" + CurrParameter.FieldName + "=*" + SearchValue + "*))";
                             break;

                             case "DoesNotEqual":
                             Filter += "(!(" + CurrParameter.FieldName + "=" + SearchValue + "))";
                             break;

                             case "GreaterThanOrEqualTo":
                             Filter += "(" + CurrParameter.FieldName + ">=" + SearchValue + ")";
                             break;
                             
                             case "LessThanOrEqualTo":
                             Filter += "(" + CurrParameter.FieldName + "<=" + SearchValue + ")"; 
                             break;

                             case "GreaterThan":
                             Filter += "(&(" + CurrParameter.FieldName + ">=" + SearchValue + ")(!(" + CurrParameter.FieldName + "=" + SearchValue + ")))"; 
                             break;

                             case "LessThan":
                             Filter += "(&(" + CurrParameter.FieldName + "<=" + SearchValue + ")(!(" + CurrParameter.FieldName + "=" + SearchValue + ")))";
                             break;

                             case "Exists":
                             Filter += "(" + CurrParameter.FieldName + "=*)";
                             break;

                             case "DoesNotExist":
                             Filter += "(!(" + CurrParameter.FieldName + "=*))";
                             break;

                             case "StartsWith":
                             Filter += "(" + CurrParameter.FieldName + "=" + SearchValue + "*)";
                             break;

                             case "DoesNotStartWith":
                             Filter += "(!(" + CurrParameter.FieldName + "=" + SearchValue + "*))";
                             break;

                             case "EndsWith":
                             Filter += "(" + CurrParameter.FieldName + "=*" + SearchValue + ")";
                             break;

                             case "DoesNotEndWith":
                             Filter += "(!(" + CurrParameter.FieldName + "=*" + SearchValue + "))";
                             break;
                         }
                        
                    }
                    Filter += ")";
                }

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
                string str = string.Empty;
                if ((BaseSearch == null) || (BaseSearch == string.Empty))
                {
                    baseLdapPath = string.Format("LDAP://{0}", (object) ADServer);
                }
                else
                {
                    baseLdapPath = string.Format("LDAP://{0}/{1}", (object) ADServer, (object) BaseSearch);
                }
                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);
                DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
                directorySearcher.Filter = "(&(objectClass=computer)" + Filter + ")";

                SearchResultCollection All = directorySearcher.FindAll();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                directorySearcher.Dispose();

                List<Computer> Results = new List<Computer>();
                if (All != null && All.Count != 0)
                {
                    foreach (SearchResult Current in All)
                    {
                        Results.Add(new Computer(Current, AdditionalAttributes));
                    }
                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }
                

                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("FoundComputers", (object) Results.ToArray());
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