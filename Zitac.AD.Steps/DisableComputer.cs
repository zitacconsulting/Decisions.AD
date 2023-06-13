using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using System;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Linq;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.CoreSteps;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Disable Computer", "Integration", "Active Directory", "Zitac", "Computer")]
    [Writable]
    public class DisableComputer : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, IDefaultInputMappingStep //, INotifyPropertyChanged
    {
     
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

        public IInputMapping[] DefaultInputs
        {
            get
            {
                IInputMapping[] inputMappingArray = new IInputMapping[1];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
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
                            
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "AD Server"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port", false, true, false));
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Computername Or DN"));
                            return dataDescriptionList.ToArray();                                              
                        }
            }
    
            public override OutcomeScenarioData[] OutcomeScenarios {
                get {

                    return new[] {
                    new OutcomeScenarioData("Done"),
                    new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")), 
                }; 
                }
            }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            int? Port = (int?)data.Data["Port"];
            string Computername = data.Data["Computername Or DN"] as string;

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
            var Filter = "(&(objectClass=computer)(|(name=" + Computername + ")(distinguishedname=" + Computername + ")))";
            try
            {

                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);
                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = LDAPHelper.GetBaseDN(connection);
                List<SearchResultEntry> ComputerResults = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "distinguishedname", "userAccountControl" }).ToList();
                string FoundComputerDN = String.Empty;
                Int64 UserAccountControl = 0;

                if (ComputerResults != null && ComputerResults.Count != 0)
                {
                    FoundComputerDN = Converters.GetStringProperty(ComputerResults[0], "distinguishedname");
                    UserAccountControl = Converters.GetIntProperty(ComputerResults[0], "userAccountControl");
                }
                else
                {
                    throw new Exception(string.Format("Unable to find computer with name or DN: '{0}' in the AD", Computername));
                }

                ModifyRequest modifyRequest = new ModifyRequest(FoundComputerDN);

                Int64 UserAccessControl = UserAccountControl;
                UserAccessControl = UserAccessControl | 0x2; 
                if (UserAccessControl != UserAccountControl)
                {
                    DirectoryAttributeModification AttributeModification = new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = "userAccountControl" };
                    AttributeModification.Add(UserAccessControl.ToString());
                    modifyRequest.Modifications.Add(AttributeModification);
                }
                else {
                    return new ResultData("Done");
                }

                ModifyResponse response = (ModifyResponse)connection.SendRequest((DirectoryRequest)modifyRequest);

                if (response.ResultCode != ResultCode.Success)
                {
                    throw new Exception("Failed to change computer attributes. Result code: " + response.ResultCode + ". Error: " + response.ErrorMessage);
                }

                connection.Dispose();
                return new ResultData("Done");

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