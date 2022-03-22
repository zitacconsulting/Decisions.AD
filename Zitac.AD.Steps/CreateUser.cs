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
    [AutoRegisterStep("Create User", "Integration", "Active Directory", "Zitac", "User")]
    [Writable]

    public class CreateUser : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged//, IDefaultInputMappingStep
    {
        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private string[] attributes;

        [PropertyClassification(8, "Use Integrated Authentication", new string[] { "Integrated Authentication" })]
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

        [PropertyClassification(10, "Attributes", new string[] { "Additional Attributes" })]
        public string[] Attributes
        {
            get
            {
                return attributes;
            }
            set
            {
                attributes = value;
                this.OnPropertyChanged("InputData");
                this.OnPropertyChanged(nameof(Attributes));
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "OU (DN)"));
                

                if (this.Attributes != null && this.Attributes.Length != 0)
                {
                    foreach (string CurrParameter in this.Attributes)
                    {
                        dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), CurrParameter){
                            Categories = new string[] {"Additional Attributes Input"}
                        });
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

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done"));
                outcomeScenarioDataList.Add(new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")));
                return outcomeScenarioDataList.ToArray();
            }
        }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            string BaseSearch = data.Data["OU (DN)"] as string;
            string UserName = data.Data["User Name"] as string;
            string[] AdditionalAttributes = data.Data["Additional Attributes"] as string[];

            string Filter = string.Empty;

            string[] ParametersList = this.Attributes;
            if (ParametersList != null && ParametersList.Length != 0)
            {

            }

            Credentials ADCredentials = new Credentials();

            if (IntegratedAuthentication)
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
                    baseLdapPath = string.Format("LDAP://{0}", (object)ADServer);
                }
                else
                {
                    baseLdapPath = string.Format("LDAP://{0}/{1}", (object)ADServer, (object)BaseSearch);
                }
                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);
                DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
                directorySearcher.Filter = "(&(objectClass=user)(objectCategory=person)" + Filter + ")";

                SearchResultCollection All = directorySearcher.FindAll();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                directorySearcher.Dispose();

                List<User> Results = new List<User>();
                if (All != null && All.Count != 0)
                {
                    foreach (SearchResult Current in All)
                    {
                        Results.Add(new User(Current, AdditionalAttributes));
                    }
                }


                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("FoundUsers", (object)Results.ToArray());
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