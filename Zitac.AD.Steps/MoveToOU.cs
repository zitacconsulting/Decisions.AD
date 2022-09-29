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
    [AutoRegisterStep("Move To OU", "Integration", "Active Directory", "Zitac")]
    [Writable]
    public class MoveToOU : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer //, INotifyPropertyChanged
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
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Account Name Or DN"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "OU DN"));
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
            string Accountname = data.Data["Account Name Or DN"] as string;
            string OUDN = data.Data["OU DN"] as string;

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
                string baseLdapPath = string.Format("LDAP://{0}", (object)ADServer);
                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);
                DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
                directorySearcher.Filter = "((|(sAMAccountName=" + Accountname + ")(distinguishedName=" + Accountname + ")))";

                SearchResult one = directorySearcher.FindOne();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                directorySearcher.Dispose();

                if (one == null)
                {
                    throw new Exception(string.Format("Unable to find object with accountname or DN: '{0}' in the AD", (object)Accountname));
                }
                DirectoryEntry childEntry = one.GetDirectoryEntry();

                
                DirectoryEntry NewOU = new DirectoryEntry("LDAP://" + OUDN, ADCredentials.ADUsername, ADCredentials.ADPassword);
                childEntry.MoveTo(NewOU);
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