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
    [AutoRegisterStep("Update User", "Integration", "Active Directory", "Zitac", "User")]
    [Writable]

    public class UpdateUser : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged, IDefaultInputMappingStep
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
                IInputMapping[] inputMappingArray = new IInputMapping[33];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Account Disabled" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Password Never Expires" };
                inputMappingArray[2] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Must Change Password On Next Login" };
                inputMappingArray[3] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Login Name (Pre-Win 2000)" };
                inputMappingArray[4] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Login Name (UPN)" };
                inputMappingArray[5] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "First Name" };
                inputMappingArray[6] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Last Name" };
                inputMappingArray[7] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Initials" };
                inputMappingArray[8] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Display Name" };
                inputMappingArray[9] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Office" };
                inputMappingArray[10] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Telephone Number" };
                inputMappingArray[11] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Description" };
                inputMappingArray[12] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Email Address" };
                inputMappingArray[13] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Web Page" };
                inputMappingArray[14] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Account Expires" };

                inputMappingArray[15] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Street" };
                inputMappingArray[16] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "PO Box" };
                inputMappingArray[17] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "City" };
                inputMappingArray[18] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "State/Province" };
                inputMappingArray[19] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Zip/Postal Code" };
                inputMappingArray[20] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Country/Region" };

                inputMappingArray[21] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Home Folder" };
                inputMappingArray[22] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Home Folder Drive Letter" };

                inputMappingArray[23] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Home Phone" };
                inputMappingArray[24] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Mobile Phone" };

                inputMappingArray[25] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Department" };
                inputMappingArray[26] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Job title" };
                inputMappingArray[27] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Company" };
                inputMappingArray[28] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Manager (DN)" };

                inputMappingArray[29] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Employee ID" };
                inputMappingArray[30] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Employee Number" };
                inputMappingArray[31] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Employee Type" };

                inputMappingArray[32] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Username Or DN"));

                // User Data
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Account Disabled") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Password Never Expires") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Must Change Password On Next Login") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Login Name (Pre-Win 2000)") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Login Name (UPN)") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "First Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Last Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Initials") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Display Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Office") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Telephone Number") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Description") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Email Address") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Web Page") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(DateTime)), "Account Expires") { Categories = new string[] { "User Data" } });

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Street") { Categories = new string[] { "User Data", "Address" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "PO Box") { Categories = new string[] { "User Data", "Address" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "City") { Categories = new string[] { "User Data", "Address" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "State/Province") { Categories = new string[] { "User Data", "Address" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Zip/Postal Code") { Categories = new string[] { "User Data", "Address" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(CountryCode)), "Country/Region") { Categories = new string[] { "User Data", "Address" } });

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Home Folder") { Categories = new string[] { "User Data", "Profile" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Home Folder Drive Letter") { Categories = new string[] { "User Data", "Profile" } });

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Home Phone") { Categories = new string[] { "User Data", "Telephones" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Mobile Phone") { Categories = new string[] { "User Data", "Telephones" } });



                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Department") { Categories = new string[] { "User Data", "Organization" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Job title") { Categories = new string[] { "User Data", "Organization" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Company") { Categories = new string[] { "User Data", "Organization" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Manager (DN)") { Categories = new string[] { "User Data", "Organization" } });


                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Employee ID") { Categories = new string[] { "User Data", "Employee" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Employee Number") { Categories = new string[] { "User Data", "Employee" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Employee Type") { Categories = new string[] { "User Data", "Employee" } });

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
            string Username = data.Data["Username Or DN"] as string;

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
            var Filter = "(&(objectClass=user)(|(sAMAccountName=" + Username + ")(distinguishedname=" + Username + ")))";
            try
            {

                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);
                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = LDAPHelper.GetBaseDN(connection);
                List<SearchResultEntry> UserResults = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "distinguishedname", "objectSid", "userAccountControl" }).ToList();
                string FoundUserDN = String.Empty;
                Int64 UserAccountControl = 0;

                if (UserResults != null && UserResults.Count != 0)
                {
                    FoundUserDN = Converters.GetStringProperty(UserResults[0], "distinguishedname");
                    UserAccountControl = Converters.GetIntProperty(UserResults[0], "userAccountControl");
                }
                else
                {
                    throw new Exception(string.Format("Unable to find user with name or DN: '{0}' in the AD", Username));
                }

                ModifyRequest modifyRequest = new ModifyRequest(FoundUserDN);

                List<AttributeValues> AttributesToAdd = new List<AttributeValues> {
                    new AttributeValues("Login Name (Pre-Win 2000)", "sAMAccountName"),
                    new AttributeValues("Login Name (UPN)", "userPrincipalName"),
                    new AttributeValues("First Name", "givenName"),
                    new AttributeValues("Last Name", "sn"),
                    new AttributeValues("Initials", "initials"),
                    new AttributeValues("Display Name", "displayName"),
                    new AttributeValues("Office", "physicalDeliveryOfficeName"),
                    new AttributeValues("Telephone Number", "telephoneNumber"),
                    new AttributeValues("Description", "description"),
                    new AttributeValues("Email Address", "mail"),
                    new AttributeValues("Web Page", "wWWHomePage"),
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
                if ((bool?)data.Data["Password Never Expires"] == true) { UserAccessControl = UserAccessControl | 0x10000; }
                if ((bool?)data.Data["Password Never Expires"] == false) { UserAccessControl = UserAccessControl & ~0x10000; }
                if (UserAccessControl != UserAccountControl)
                {
                    DirectoryAttributeModification AttributeModification = new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = "userAccountControl" };
                    AttributeModification.Add(UserAccessControl.ToString());
                    modifyRequest.Modifications.Add(AttributeModification);
                }

                if ((bool?)data.Data["Must Change Password On Next Login"] == true)
                {
                    DirectoryAttributeModification AttributeModification = new DirectoryAttributeModification { Operation = DirectoryAttributeOperation.Replace, Name = "pwdLastSet" };
                    AttributeModification.Add((string)"0");
                    modifyRequest.Modifications.Add(AttributeModification);
                }

                ModifyResponse response = (ModifyResponse)connection.SendRequest((DirectoryRequest)modifyRequest);

                if (response.ResultCode != ResultCode.Success)
                {
                    throw new Exception("Failed to change user attributes. Result code: " + response.ResultCode + ". Error: " + response.ErrorMessage);
                }

                return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "DN", (object)FoundUserDN } });




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