using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using System;
using System.Text;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Linq;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.Mapping.InputImpl;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.ComponentModel;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Create User", "Integration", "Active Directory", "Zitac", "User")]
    [Writable]

    public class CreateUser : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, INotifyPropertyChanged, IDefaultInputMappingStep
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
                IInputMapping[] inputMappingArray = new IInputMapping[31];
                inputMappingArray[0] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Account Disabled" };
                inputMappingArray[1] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Password Never Expires" };
                inputMappingArray[2] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Must Change Password On Next Login" };
                inputMappingArray[3] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Initials" };
                inputMappingArray[4] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Display Name" };
                inputMappingArray[5] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Office" };
                inputMappingArray[6] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Telephone Number" };
                inputMappingArray[7] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Description" };
                inputMappingArray[8] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Email Address" };
                inputMappingArray[9] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Web Page" };
                inputMappingArray[10] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Account Expires" };

                inputMappingArray[11] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Street" };
                inputMappingArray[12] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "PO Box" };
                inputMappingArray[13] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "City" };
                inputMappingArray[14] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "State/Province" };
                inputMappingArray[15] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Zip/Postal Code" };
                inputMappingArray[16] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Country/Region" };

                inputMappingArray[17] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Home Folder" };
                inputMappingArray[18] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Home Folder Drive Letter" };

                inputMappingArray[19] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Home Phone" };
                inputMappingArray[20] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Mobile Phone" };

                inputMappingArray[21] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Department" };
                inputMappingArray[22] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Job title" };
                inputMappingArray[23] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Company" };
                inputMappingArray[24] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Manager (DN)" };

                inputMappingArray[25] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Employee ID" };
                inputMappingArray[26] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Employee Number" };
                inputMappingArray[27] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Employee Type" };

                inputMappingArray[28] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Port" };
                inputMappingArray[29] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "First Name" };
                inputMappingArray[30] = (IInputMapping)new IgnoreInputMapping() { InputDataName = "Last Name" };
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port",false, true, false));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "OU (DN)"));

                // User Data
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Account Disabled") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Password Never Expires") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Must Change Password On Next Login") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Login Name (Pre-Win 2000)") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Login Name (UPN)") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Common Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "First Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Last Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Initials") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Display Name") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Office") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Telephone Number") { Categories = new string[] { "User Data" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Password") { Categories = new string[] { "User Data" }, EditorAttribute = (PropertyEditorAttribute)new PasswordTextAttribute() });
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
            string OU = data.Data["OU (DN)"] as string;

            string sAMAccountName = data.Data["Login Name (Pre-Win 2000)"] as string;
            string Passwd = data.Data["Password"] as string;

            int UserAccessControl = 512;
            if ((bool?)data.Data["Account Disabled"] == true) { UserAccessControl = UserAccessControl | 0x2; }
            if ((bool?)data.Data["Password Never Expires"] == true) { UserAccessControl = UserAccessControl | 0x10000; }


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

            try
            {

                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);
                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);

                List<DirectoryAttribute> attributes = new List<DirectoryAttribute>();
                attributes.Add(new DirectoryAttribute("sAMAccountName", sAMAccountName));
                attributes.Add(new DirectoryAttribute("unicodePwd", Encoding.Unicode.GetBytes($"\"{Passwd}\"")));
                attributes.Add(new DirectoryAttribute("objectClass", "user"));

                if (data.Data["Login Name (UPN)"] != null && data.Data["Login Name (UPN)"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("userPrincipalName", (string)data.Data["Login Name (UPN)"])); }
                if (data.Data["First Name"] != null && data.Data["First Name"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("givenName", (string)data.Data["First Name"])); }
                if (data.Data["Last Name"] != null && data.Data["Last Name"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("sn", (string)data.Data["Last Name"])); }

        
                attributes.Add(new DirectoryAttribute("userAccountControl", UserAccessControl.ToString()));
                if (data.Data["Initials"] != null && data.Data["Initials"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("initials", (string)data.Data["Initials"])); }
                if (data.Data["Display Name"] != null && data.Data["Display Name"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("displayName", (string)data.Data["Display Name"])); }
                if (data.Data["Office"] != null && data.Data["Office"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("physicalDeliveryOfficeName", (string)data.Data["Office"])); }
                if (data.Data["Telephone Number"] != null && data.Data["Telephone Number"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("telephoneNumber", (string)data.Data["Telephone Number"])); }
                if (data.Data["Description"] != null && data.Data["Description"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("description", (string)data.Data["Description"])); }
                if (data.Data["Email Address"] != null && data.Data["Email Address"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("mail", (string)data.Data["Email Address"])); }
                if (data.Data["Web Page"] != null && data.Data["Web Page"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("wWWHomePage", (string)data.Data["Web Page"])); }
                if (data.Data["Account Expires"] != null && data.Data["Account Expires"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("accountExpires", Convert.ToString(((DateTime)data.Data["Account Expires"]).ToFileTimeUtc()))); }

                if (data.Data["Street"] != null && data.Data["Street"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("streetAddress", (string)data.Data["Street"])); }
                if (data.Data["PO Box"] != null && data.Data["PO Box"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("postOfficeBox", (string)data.Data["PO Box"])); }
                if (data.Data["City"] != null && data.Data["City"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("l", (string)data.Data["City"])); }
                if (data.Data["State/Province"] != null && data.Data["State/Province"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("st", (string)data.Data["State/Province"])); }
                if (data.Data["Zip/Postal Code"] != null && data.Data["Zip/Postal Code"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("postalCode", (string)data.Data["Zip/Postal Code"])); }
                if (data.Data["Country/Region"] != null && data.Data["Country/Region"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("c", (string)data.Data["Country/Region"].ToString())); }

                if (data.Data["Home Folder"] != null && data.Data["Home Folder"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("homeDirectory", (string)data.Data["Home Folder"])); }
                if (data.Data["Home Folder Drive Letter"] != null && data.Data["Home Folder Drive Letter"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("homeDrive", (string)data.Data["Home Folder Drive Letter"])); }

                if (data.Data["Home Phone"] != null && data.Data["Home Phone"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("homePhone", (string)data.Data["Home Phone"])); }
                if (data.Data["Mobile Phone"] != null && data.Data["Mobile Phone"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("mobile", (string)data.Data["Mobile Phone"])); }

                if (data.Data["Department"] != null && data.Data["Department"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("department", (string)data.Data["Department"])); }
                if (data.Data["Job title"] != null && data.Data["Job title"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("title", (string)data.Data["Job title"])); }
                if (data.Data["Company"] != null && data.Data["Company"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("company", (string)data.Data["Company"])); }
                if (data.Data["Manager (DN)"] != null && data.Data["Manager (DN)"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("manager", (string)data.Data["Manager (DN)"])); }

                if (data.Data["Employee ID"] != null && data.Data["Employee ID"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("employeeID", (string)data.Data["Employee ID"])); }
                if (data.Data["Employee Number"] != null && data.Data["Employee Number"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("employeeNumber", (string)data.Data["Employee Number"])); }
                if (data.Data["Employee Type"] != null && data.Data["Employee Type"].ToString().Length != 0) { attributes.Add(new DirectoryAttribute("employeeType", (string)data.Data["Employee Type"])); }
                
                if ((bool?)data.Data["Must Change Password On Next Login"] == true) { attributes.Add(new DirectoryAttribute("pwdLastSet", "0")); }
                
                string[] ParametersList = this.Attributes;
                if (ParametersList != null && ParametersList.Length != 0)
                {
                    foreach (string CurrParameter in ParametersList)
                    {
                        if (data.Data[CurrParameter] != null && data.Data[CurrParameter].ToString().Length != 0) { attributes.Add(new DirectoryAttribute(CurrParameter, (string)data.Data[CurrParameter])); }
                    }
                }

                string newUserDN = $"CN={data.Data["Common Name"]},{OU}";
                Console.WriteLine(newUserDN);
                AddRequest addRequest = new AddRequest(newUserDN, attributes.ToArray());
                AddResponse addResponse = (AddResponse)connection.SendRequest(addRequest);
                if (addResponse.ResultCode != ResultCode.Success)
                {
                    throw new Exception($"User creation failed. ResultCode: {addResponse.ResultCode}");
                }

 
                return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "DN", (object)newUserDN } });




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