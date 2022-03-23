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

                // User Data
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Account Disabled") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Password Never Expires") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(bool?)), "Must Change Password On Next Login") { Categories = new string[] { "User Data", "Flags" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "sAMAccountName") { Categories = new string[] { "User Data" } });
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Country/Region") { Categories = new string[] { "User Data", "Address" } });
                
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Home Folder") { Categories = new string[] { "User Data", "Profile" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Home Folder Drive Letter") { Categories = new string[] { "User Data", "Profile" } });

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Home Phone") { Categories = new string[] { "User Data", "Telephones" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Mobile Phone") { Categories = new string[] { "User Data", "Telephones" } });



                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Department") { Categories = new string[] { "User Data", "Organization" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Job title") { Categories = new string[] { "User Data", "Organization" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Company") { Categories = new string[] { "User Data", "Organization" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Manager (DN)") { Categories = new string[] { "User Data", "Organization" } });


                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Empolyee ID") { Categories = new string[] { "User Data", "Employee" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Empolyee Number") { Categories = new string[] { "User Data", "Employee" } });
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Empolyee Type") { Categories = new string[] { "User Data", "Employee" } });
                
                
                


                //Additional Attributes
                if (this.Attributes != null && this.Attributes.Length != 0)
                {
                    foreach (string CurrParameter in this.Attributes)
                    {
                        dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), CurrParameter)
                        {
                            Categories = new string[] { "Additional Attributes Input" }
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
            string OU = data.Data["OU (DN)"] as string;

            string FirstName = data.Data["First Name"] as string;
            string LastName = data.Data["Last Name"] as string;
            string sAMAccountName = data.Data["sAMAccountName"] as string;
            string Passwd = data.Data["Password"] as string;

            int UserAccessControl = 512;
                if((bool?)data.Data["Account Disabled"] == true) {UserAccessControl = UserAccessControl | 0x2;}
                if((bool?)data.Data["Password Never Expires"] == true) {UserAccessControl = UserAccessControl | 0x10000;}


            string[] AdditionalAttributes = data.Data["Additional Attributes"] as string[];

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


                string baseLdapPath = string.Format("LDAP://{0}/{1}", (object)ADServer, (object)OU);

                DirectoryEntry ouEntry = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);

                DirectoryEntry childEntry = ouEntry.Children.Add("CN=" + FirstName + " " + LastName, "user");
                childEntry.Properties["sAMAccountName"].Value = sAMAccountName;
                childEntry.Properties["givenName"].Value = FirstName;
                childEntry.Properties["sn"].Value = LastName;
                childEntry.CommitChanges();
                ouEntry.CommitChanges();
                childEntry.Invoke("SetPassword", new object[] { Passwd });
                childEntry.CommitChanges();
                if(UserAccessControl != 0){ childEntry.Properties["userAccountControl"].Value = UserAccessControl; }
                if(data.Data["Initials"] != null){ childEntry.Properties["initials"].Value = (string)data.Data["Initials"];}
                if(data.Data["Display Name"] != null){ childEntry.Properties["displayName"].Value = (string)data.Data["Display Name"];}
                if(data.Data["Office"] != null){ childEntry.Properties["physicalDeliveryOfficeName"].Value = (string)data.Data["Office"];}
                if(data.Data["Telephone Number"] != null){ childEntry.Properties["telephoneNumber"].Value = (string)data.Data["Telephone Number"];}
                if(data.Data["Description"] != null){ childEntry.Properties["description"].Value = (string)data.Data["Description"];}
                if(data.Data["Email Address"] != null){ childEntry.Properties["mail"].Value = (string)data.Data["Email Address"];}
                if(data.Data["Web Page"] != null){ childEntry.Properties["wWWHomePage"].Value = (string)data.Data["Web Page"];}
                if(data.Data["Account Expires"] != null){ childEntry.Properties["accountExpires"].Value = Convert.ToString(((DateTime)data.Data["Account Expires"]).ToFileTimeUtc());}

                if(data.Data["Street"] != null){ childEntry.Properties["streetAddress"].Value = (string)data.Data["Street"];}
                if(data.Data["PO Box"] != null){ childEntry.Properties["postOfficeBox"].Value = (string)data.Data["PO Box"];}
                if(data.Data["City"] != null){ childEntry.Properties["l"].Value = (string)data.Data["City"];}
                if(data.Data["State/Province"] != null){ childEntry.Properties["st"].Value = (string)data.Data["State/Province"];}
                if(data.Data["Zip/Postal Code"] != null){ childEntry.Properties["postalCode"].Value = (string)data.Data["Zip/Postal Code"];}
                if(data.Data["Country/Region"] != null){ childEntry.Properties["c"].Value = (string)data.Data["Country/Region"];}
                
                if(data.Data["Home Folder"] != null){ childEntry.Properties["homeDirectory"].Value = (string)data.Data["Home Folder"];}
                if(data.Data["Home Folder Drive Letter"] != null){ childEntry.Properties["homeDrive"].Value = (string)data.Data["Home Folder Drive Letter"];}

                if(data.Data["Home Phone"] != null){ childEntry.Properties["homePhone"].Value = (string)data.Data["Home Phone"];}
                if(data.Data["Mobile Phone"] != null){ childEntry.Properties["mobile"].Value = (string)data.Data["Mobile Phone"];}

                if(data.Data["Department"] != null){ childEntry.Properties["department"].Value = (string)data.Data["Department"];}
                if(data.Data["Job title"] != null){ childEntry.Properties["title"].Value = (string)data.Data["Job title"];}
                if(data.Data["Company"] != null){ childEntry.Properties["company"].Value = (string)data.Data["Company"];}
                if(data.Data["Manager (DN)"] != null){ childEntry.Properties["manager"].Value = (string)data.Data["Manager (DN)"];}

                if(data.Data["Empolyee ID"] != null){ childEntry.Properties["employeeID"].Value = (string)data.Data["Empolyee ID"];}
                if(data.Data["Empolyee Number"] != null){ childEntry.Properties["employeeNumber"].Value = (string)data.Data["Empolyee Number"];}
                if(data.Data["Empolyee Type"] != null){ childEntry.Properties["employeeType"].Value = (string)data.Data["Empolyee Type"];}


                if((bool?)data.Data["Must Change Password On Next Login"] == true) {childEntry.Properties["pwdLastSet"].Value = 0;}
                childEntry.CommitChanges();

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
    }
}