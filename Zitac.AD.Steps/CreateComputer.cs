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
using System.Text;
namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Create Computer", "Integration", "Active Directory", "Zitac", "Computer")]
    [Writable]

    public class CreateComputer : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged, IDefaultInputMappingStep
    {
        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private string[] attributes;

        [PropertyClassification(2, "Use Integrated Authentication", new string[] { "Integrated Authentication" })]
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

        public IInputMapping[] DefaultInputs
        {
            get
            {
                IInputMapping[] inputMappingArray = new IInputMapping[4];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Account Disabled" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Description" };
                inputMappingArray[2] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Location" };
                inputMappingArray[3] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Managed By (DN)" };




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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "OU (DN)"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Computer Name"));

                // Computer Data
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Account Disabled") { Categories = new string[] { "General" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Description") { Categories = new string[] { "General" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Location") { Categories = new string[] { "General" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Managed By (DN)") { Categories = new string[] { "General" } });



                //Additional Attributes
                if (this.Attributes != null && this.Attributes.Length != 0)
                {
                    foreach (string CurrParameter in this.Attributes)
                    {
                        dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), CurrParameter)
                        {
                            Categories = new string[] { "Additional Attributes Input" },
                            SortIndex = 3
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

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(string), "DN")));
                outcomeScenarioDataList.Add(new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")));
                return outcomeScenarioDataList.ToArray();
            }
        }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            string ComputerName = data.Data["Computer Name"] as string;
            string OU = data.Data["OU (DN)"] as string;

            int UserAccessControl = 512;
            if ((bool?)data.Data["Account Disabled"] == true) { UserAccessControl = UserAccessControl | 0x2; }


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



            string baseLdapPath = string.Format("LDAP://{0}/{1}", (object)ADServer, (object)OU);
            try
            {
                DirectoryEntry ouEntry = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);

                DirectoryEntry childEntry = ouEntry.Children.Add("CN=" + ComputerName, "computer");
                childEntry.Properties["sAMAccountName"].Value = (ComputerName + "$");
                childEntry.CommitChanges();


                try
                {

                    //if (UserAccessControl != 0) { childEntry.Properties["userAccountControl"].Value = UserAccessControl; }
                    if (data.Data["Description"] != null && (data.Data["Description"]).ToString().Length != 0) { childEntry.Properties["description"].Value = (string)data.Data["Description"]; }
                    if (data.Data.ContainsKey("Description") && (data.Data["Description"]) == null) { childEntry.Properties["description"].Clear(); }

                    if (data.Data["Location"] != null && (data.Data["Location"]).ToString().Length != 0) { childEntry.Properties["location"].Value = (string)data.Data["Location"]; }
                    if (data.Data.ContainsKey("Location") && (data.Data["Location"]) == null) { childEntry.Properties["location"].Clear(); }

                    if (data.Data["Managed By (DN)"] != null && (data.Data["Managed By (DN)"]).ToString().Length != 0) { childEntry.Properties["managedBy"].Value = (string)data.Data["Managed By (DN)"]; }
                    if (data.Data.ContainsKey("Managed By (DN)") && (data.Data["Managed By (DN)"]) == null) { childEntry.Properties["managedBy"].Clear(); }

                    childEntry.CommitChanges();

                    if ((bool?)data.Data["Account Disabled"] == true) {
                        childEntry.InvokeSet("AccountDisabled", true);
                      }
                    else {
                        childEntry.InvokeSet("AccountDisabled", false);
                    }


                    string[] ParametersList = this.Attributes;
                    if (ParametersList != null && ParametersList.Length != 0)
                    {
                        foreach (string CurrParameter in ParametersList)
                        {
                            if (data.Data[CurrParameter] != null && (data.Data[CurrParameter]).ToString().Length != 0) { childEntry.Properties[CurrParameter].Value = (string)data.Data[CurrParameter]; }
                            if (data.Data.ContainsKey(CurrParameter) && (data.Data[CurrParameter]) == null) { childEntry.Properties[CurrParameter].Clear(); }
                        }
                        childEntry.CommitChanges();
                    }

                    return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "DN", (object)childEntry.Properties["distinguishedName"].Value } });




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