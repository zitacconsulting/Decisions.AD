using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
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
    [AutoRegisterStep("Create Computer", "Integration", "Active Directory", "Zitac", "Computer")]
    [Writable]

    public class CreateComputer : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged, IDefaultInputMappingStep
    {
        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool useSSL = true;

        [WritableValue]
        private bool ignoreInvalidCert;

        [WritableValue]
        private string[] attributes;

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
                IInputMapping[] inputMappingArray = new IInputMapping[5];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Account Disabled" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Description" };
                inputMappingArray[2] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Location" };
                inputMappingArray[3] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Managed By (DN)" };
                inputMappingArray[4] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };




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
            int? Port = (int?)data.Data["Port"];
            string ComputerName = data.Data["Computer Name"] as string;
            string OU = data.Data["OU (DN)"] as string;

            int UserAccessControl = 4096;
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

            try
            {

                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials.ADUsername, ADCredentials.ADPassword, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);
                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);

                List<DirectoryAttribute> attributes = new List<DirectoryAttribute>();
                attributes.Add(new DirectoryAttribute("sAMAccountName", (ComputerName + "$")));
                attributes.Add(new DirectoryAttribute("objectClass", "computer"));

                attributes.Add(new DirectoryAttribute("userAccountControl", UserAccessControl.ToString()));

                if (data.Data["Description"] != null && data.Data["Description"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("description", (string)data.Data["Description"])); }
                if (data.Data["Location"] != null && data.Data["Location"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("location", (string)data.Data["Location"])); }
                if (data.Data["Managed By (DN)"] != null && data.Data["Managed By (DN)"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("managedBy", (string)data.Data["Managed By (DN)"])); }

                string[] ParametersList = this.Attributes;
                if (ParametersList != null && ParametersList.Length != 0)
                {
                    foreach (string CurrParameter in ParametersList)
                    {
                        if (data.Data[CurrParameter] != null && data.Data[CurrParameter].ToString().Length != 0) { attributes.Add(new DirectoryAttribute(CurrParameter, (string)data.Data[CurrParameter])); }
                    }
                }

                string newComputerDN = $"CN={ComputerName},{OU}";
                Console.WriteLine(newComputerDN);
                AddRequest addRequest = new AddRequest(newComputerDN, attributes.ToArray());
                AddResponse addResponse = (AddResponse)connection.SendRequest(addRequest);
                if (addResponse.ResultCode != ResultCode.Success)
                {
                    throw new Exception($"Computer creation failed. ResultCode: {addResponse.ResultCode}");
                }

                return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "DN", (object)newComputerDN } });

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