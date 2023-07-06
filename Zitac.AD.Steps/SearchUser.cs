using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.ComponentModel;
using System.DirectoryServices.Protocols;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Search User", "Integration", "Active Directory", "Zitac", "User")]
    [Writable]
    public class SearchUser : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged, IDefaultInputMappingStep
    {

        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool useSSL = true;

        [WritableValue]
        private bool ignoreInvalidCert;

        [WritableValue]
        private bool combineFiltersUsingAnd;

        [WritableValue]
        private bool showOutcomeforNoResults;

        [WritableValue]
        private bool? passwordNeverExpires;

        [WritableValue]
        private bool? accountEnabled;



        [WritableValue]
        private SearchParameters[] qParams;

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


        [PropertyClassification(9, "Combine Filters Using And", new string[] { "Search Definition" })]
        public bool CombineFiltersUsingAnd
        {
            get { return combineFiltersUsingAnd; }
            set { combineFiltersUsingAnd = value; }

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

        [PropertyClassification(10, "Search Criteria", new string[] { "Search Definition" })]
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
                this.OnPropertyChanged(nameof(QueryParams));
            }
        }

        [PropertyClassification(11, "Password Never Expires", new string[] { "Search Definition" })]
        public bool? PasswordNeverExpires
        {
            get { return passwordNeverExpires; }
            set { passwordNeverExpires = value; }

        }

        [PropertyClassification(11, "Account Enabled", new string[] { "Search Definition" })]
        public bool? AccountEnabled
        {
            get { return accountEnabled; }
            set { accountEnabled = value; }

        }
        public IInputMapping[] DefaultInputs
        {
            get
            {
                IInputMapping[] inputMappingArray = new IInputMapping[4];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Search Base (DN)" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Additional Attributes" };
                inputMappingArray[2] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
                inputMappingArray[3] = (IInputMapping)new ConstantInputMapping() { InputDataName = "Scope", Value = SearchScope.Subtree };
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
            get
            {

                List<DataDescription> dataDescriptionList = new List<DataDescription>();
                if (!IntegratedAuthentication)
                {
                    dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(Credentials)), "Credentials"));
                }

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "AD Server"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port", false, true, false));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(SearchScope)), "Scope", false, true, true));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Search Base (DN)"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Additional Attributes", true, true, true));

                SearchParameters[] ParametersList = this.GetSearchParameters();
                if (ParametersList != null && ParametersList.Length != 0)
                {
                    foreach (SearchParameters CurrParameter in ParametersList)
                    {
                        if (!new[] { "Exists", "DoesNotExist" }.Contains(CurrParameter.MatchCriteria))
                        {
                            if (CurrParameter.DataType == "Date")
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(DateTime)), CurrParameter.Alias));
                            }
                            else if (CurrParameter.DataType == "Int32")
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(Int32)), CurrParameter.Alias));
                            }
                            else if (CurrParameter.DataType == "Int64")
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(Int64)), CurrParameter.Alias));
                            }
                            else
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), CurrParameter.Alias));
                            }
                        }
                    }
                }

                return dataDescriptionList.ToArray();
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(User), "FoundUsers", true)));
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
            string BaseSearch = data.Data["Search Base (DN)"] as string;
            int? Port = (int?)data.Data["Port"];
            SearchScope Scope = (SearchScope)data.Data["Scope"];
            string UserName = data.Data["User Name"] as string;
            List<string> AdditionalAttributes = (data.Data["Additional Attributes"] as string[])?.ToList();

            string Filter = string.Empty;

            SearchParameters[] ParametersList = this.GetSearchParameters();
            if (ParametersList != null && ParametersList.Length != 0)
            {
                if (CombineFiltersUsingAnd)
                {
                    Filter = "(&";
                }
                else
                {
                    Filter = "(|";
                }
                foreach (SearchParameters CurrParameter in ParametersList)
                {
                    object ParameterValue = data.Data[CurrParameter.Alias] as object;

                    string Type = ParameterValue.GetType().ToString();
                    string TextSearchValue;

                    if (Type == "System.DateTime")
                    {
                        DateTime SearchValue = (DateTime)ParameterValue;
                        TextSearchValue = SearchValue.ToFileTime().ToString();
                    }
                    else
                    {
                        TextSearchValue = ParameterValue.ToString();
                    }

                    switch (CurrParameter.MatchCriteria)
                    {
                        case "Equals":
                            Filter += "(" + CurrParameter.FieldName + "=" + TextSearchValue + ")";
                            break;

                        case "Contains":
                            Filter += "(" + CurrParameter.FieldName + "=*" + TextSearchValue + "*)";
                            break;

                        case "DoesNotContain":
                            Filter += "(!(" + CurrParameter.FieldName + "=*" + TextSearchValue + "*))";
                            break;

                        case "DoesNotEqual":
                            Filter += "(!(" + CurrParameter.FieldName + "=" + TextSearchValue + "))";
                            break;

                        case "GreaterThanOrEqualTo":
                            Filter += "(" + CurrParameter.FieldName + ">=" + TextSearchValue + ")";
                            break;

                        case "LessThanOrEqualTo":
                            Filter += "(" + CurrParameter.FieldName + "<=" + TextSearchValue + ")";
                            break;

                        case "GreaterThan":
                            Filter += "(&(" + CurrParameter.FieldName + ">=" + TextSearchValue + ")(!(" + CurrParameter.FieldName + "=" + TextSearchValue + ")))";
                            break;

                        case "LessThan":
                            Filter += "(&(" + CurrParameter.FieldName + "<=" + TextSearchValue + ")(!(" + CurrParameter.FieldName + "=" + TextSearchValue + ")))";
                            break;

                        case "Exists":
                            Filter += "(" + CurrParameter.FieldName + "=*)";
                            break;

                        case "DoesNotExist":
                            Filter += "(!(" + CurrParameter.FieldName + "=*))";
                            break;

                        case "StartsWith":
                            Filter += "(" + CurrParameter.FieldName + "=" + TextSearchValue + "*)";
                            break;

                        case "DoesNotStartWith":
                            Filter += "(!(" + CurrParameter.FieldName + "=" + TextSearchValue + "*))";
                            break;

                        case "EndsWith":
                            Filter += "(" + CurrParameter.FieldName + "=*" + TextSearchValue + ")";
                            break;

                        case "DoesNotEndWith":
                            Filter += "(!(" + CurrParameter.FieldName + "=*" + TextSearchValue + "))";
                            break;
                    }

                }
                Filter += ")";
            }
            if (PasswordNeverExpires != null || AccountEnabled != null)
            {
                Filter = "(&" + Filter;
                if (PasswordNeverExpires == true) { Filter += "(userAccountControl:1.2.840.113556.1.4.803:=65536)"; }
                if (PasswordNeverExpires == false) { Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=65536))"; }
                if (AccountEnabled == true) { Filter += "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"; }
                if (AccountEnabled == false) { Filter += "(userAccountControl:1.2.840.113556.1.4.803:=2)"; }
                Filter += ")";
            }

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

            List<string> BaseAttributeList = User.UserAttributes;

            if (AdditionalAttributes != null)
            {
                BaseAttributeList.AddRange(AdditionalAttributes);
            }
            else
            {
                AdditionalAttributes = new List<string>();
            }

            Filter = "(&(objectClass=user)(objectCategory=person)" + Filter + ")";
            Console.WriteLine(Filter);

            try
            {
                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);


                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = string.Empty;
                if (String.IsNullOrEmpty(BaseSearch))
                {
                    BaseDN = LDAPHelper.GetBaseDN(connection);
                }
                else
                {
                    BaseDN = BaseSearch;
                }
                List<SearchResultEntry> Results = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, Scope, Filter, BaseAttributeList).ToList();

                List<User> Users = new List<User>();

                if (Results != null && Results.Count != 0)
                {

                    int MaxPasswordAge = LDAPHelper.GetADPasswordExpirationPolicy(connection, null);
                    foreach (SearchResultEntry User in Results)
                    {

                        Users.Add(new(User, AdditionalAttributes, MaxPasswordAge));
                    }

                }
                else if (ShowOutcomeforNoResults)
                {
                    return new ResultData("No Results");
                }


                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("FoundUsers", (object)Users.ToArray());
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