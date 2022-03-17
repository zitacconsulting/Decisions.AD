using ActiveDirectory;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Service.Debugging.DebugData;
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
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.ComponentModel;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Change Password", "Integration", "Active Directory", "Zitac")]
    [Writable]
    public class ChangePassword : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer //, INotifyPropertyChanged
    {
     
        [WritableValue]
        private bool integratedAuthentication;

        [PropertyClassification(new string[]{"Integrated Authentication"})]
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

            public DataDescription[] InputData
            {
                    get {
                        
                        List<DataDescription> dataDescriptionList = new List<DataDescription>();
                            if(!IntegratedAuthentication)
                            {
                                dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (Credentials)), "Credentials"));
                            }
                            
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "AD Server"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Username"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "New Password")
                                {
                                EditorAttribute = (PropertyEditorAttribute) new PasswordTextAttribute()
                                });
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
            string Username = data.Data["Username"] as string;
            string Password = data.Data["New Password"] as string;

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
                string distinguishedName = GetDistinguishedName.GetObjectDistinguishedName(GetDistinguishedName.ADObjectType.User, Username, ADCredentials.ADUsername, ADCredentials.ADPassword, ADServer);
                if (string.IsNullOrEmpty(distinguishedName))
                    throw new Exception(string.Format("Unable to find user with login name: '{0}' in the AD", (object) Username));
                DirectoryEntry directoryEntry = new DirectoryEntry(distinguishedName, ADCredentials.ADUsername, ADCredentials.ADPassword);
                directoryEntry.Invoke("SetPassword", new object[] { Password });

                directoryEntry.CommitChanges();
                directoryEntry.Close();
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
                return new ResultData("Done");
        }
    }

}