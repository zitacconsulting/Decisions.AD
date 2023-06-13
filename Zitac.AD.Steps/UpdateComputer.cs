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
    [AutoRegisterStep("Update Computer", "Integration", "Active Directory", "Zitac", "Computer")]
    [Writable]

    public class UpdateComputer : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged, IDefaultInputMappingStep
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Computer Name Or DN"));

                // User Data
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
            string ComputerName = data.Data["Computer Name Or DN"] as string;

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

            var Filter = "(&(objectClass=computer)(|(name=" + ComputerName + ")(distinguishedName=" + ComputerName + ")))";
            try
            {

                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);
                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = LDAPHelper.GetBaseDN(connection);
                List<SearchResultEntry> UserResults = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "distinguishedname", "userAccountControl" }).ToList();
                string FoundComputerDN = String.Empty;
                Int64 UserAccountControl = 0;

                if (UserResults != null && UserResults.Count != 0)
                {
                    FoundComputerDN = Converters.GetStringProperty(UserResults[0], "distinguishedname");
                    UserAccountControl = Converters.GetIntProperty(UserResults[0], "userAccountControl");
                }
                else
                {
                    throw new Exception(string.Format("Unable to find computer with name or DN: '{0}' in the AD", ComputerName));
                }

                ModifyRequest modifyRequest = new ModifyRequest(FoundComputerDN);

                List<AttributeValues> AttributesToAdd = new List<AttributeValues> {
                    new AttributeValues("Description", "description"),
                    new AttributeValues("Location", "location"),
                    new AttributeValues("Managed By (DN)", "managedBy"),
                    new AttributeValues("Account Expires", "accountExpires"),
                    new AttributeValues("Street", "streetAddress"),
                    new AttributeValues("PO Box", "postOfficeBox"),
                    new AttributeValues("City", "l"),
                    new AttributeValues("State/Province", "st"),
                    new AttributeValues("Zip/Postal Code", "postalCode"),
                    new AttributeValues("Country/Region", "c"),
                    new AttributeValues("Home Folder", "homeDirectory"),
                    new AttributeValues("Home Folder Drive Letter", "homeDrive"),
                    new AttributeValues("Home Phone", "homePhone"),
                    new AttributeValues("Mobile Phone", "mobile"),
                    new AttributeValues("Department", "department"),
                    new AttributeValues("Job title", "title"),
                    new AttributeValues("Company", "company"),
                    new AttributeValues("Manager (DN)", "manager"),
                    new AttributeValues("Employee ID", "employeeID"),
                    new AttributeValues("Employee Number", "employeeNumber"),
                    new AttributeValues("Employee Type", "employeeType")
                };

                string[] ParametersList = this.Attributes;
                if (ParametersList != null && ParametersList.Length != 0)
                {
                    foreach (string CurrParameter in ParametersList)
                    {
                        AttributesToAdd.Add(new AttributeValues(CurrParameter,CurrParameter));
                    }
                }

                foreach (AttributeValues Attribute in AttributesToAdd)
                {
                    DirectoryAttributeModification ToAddAttribute = LDAPHelper.CreateAttributeModification(data, Attribute);
                    if (ToAddAttribute != null) { modifyRequest.Modifications.Add(ToAddAttribute); }
                }

                Int64 UserAccessControl = UserAccountControl;
                if ((bool?)data.Data["Account Disabled"] == true) { UserAccessControl = UserAccessControl | 0x2; }
                if ((bool?)data.Data["Account Disabled"] == false) { UserAccessControl = UserAccessControl & ~0x2; }
                if (UserAccessControl != UserAccountControl)
                {
                    DirectoryAttributeModification AttributeModification = new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = "userAccountControl" };
                    AttributeModification.Add(UserAccessControl.ToString());
                    modifyRequest.Modifications.Add(AttributeModification);
                }

                ModifyResponse response = (ModifyResponse)connection.SendRequest((DirectoryRequest)modifyRequest);

                if (response.ResultCode != ResultCode.Success)
                {
                    throw new Exception("Failed to change computer attributes. Result code: " + response.ResultCode + ". Error: " + response.ErrorMessage);
                }

                return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "DN", (object)FoundComputerDN } });

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