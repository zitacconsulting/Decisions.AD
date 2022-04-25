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
using System.Security.Principal;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Add To Group", "Integration", "Active Directory", "Zitac", "Group")]
    [Writable]
    public class AddToGroup : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer //, INotifyPropertyChanged
    {

        [WritableValue]
        private bool integratedAuthentication;

        [PropertyClassification(new string[] { "Integrated Authentication" })]
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
            get
            {

                List<DataDescription> dataDescriptionList = new List<DataDescription>();
                if (!IntegratedAuthentication)
                {
                    dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(Credentials)), "Credentials"));
                }

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "AD Server"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Group Name or DN"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Object Account Name or DN"));
                return dataDescriptionList.ToArray();
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {

                return new[] {
                    new OutcomeScenarioData("Done"),
                    new OutcomeScenarioData("Already Member",new DataDescription(typeof(string), "SID")),
                    new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")),
                };
            }
        }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            string Group = data.Data["Group Name or DN"] as string;
            string Object = data.Data["Object Account Name or DN"] as string;

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

                string baseLdapPath = string.Format("LDAP://{0}", (object)ADServer);
                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);
                DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
                directorySearcher.Filter = "(&(objectCategory=group)(objectClass=group)(|(sAMAccountName=" + Group + ")(distinguishedName=" + Group + ")))";

                SearchResult FoundGroup = directorySearcher.FindOne();

                DirectorySearcher directorySearcherObject = new DirectorySearcher(searchRoot);
                directorySearcherObject.Filter = "(|(sAMAccountName=" + Object + ")(distinguishedName=" + Object + "))";

                SearchResult FoundObject = directorySearcherObject.FindOne();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                directorySearcher.Dispose();

                if (FoundGroup == null)
                {
                    throw new Exception(string.Format("Unable to find group with name or DN: '{0}' in the AD", (object)Group));
                }

                if (FoundObject == null)
                {
                    throw new Exception(string.Format("Unable to find Object with name or DN: '{0}' in the AD", (object)Object));
                }

                DirectoryEntry ent = FoundGroup.GetDirectoryEntry();
                DirectoryEntry Obj = FoundObject.GetDirectoryEntry();


                string objectID = (new SecurityIdentifier((byte[])ent.Properties["objectSid"][0],0)).ToString();
                string GroupSid = objectID.Split("-").Last();

                if(GroupSid == Obj.Properties["primaryGroupID"].Value.ToString()) {return new ResultData("Already Member");}
                PropertyValueCollection groups = Obj.Properties["memberOf"];
                foreach (string g in groups)
                {

                    if (g.Equals(ent.Properties["distinguishedName"].Value))
                    {
                        return new ResultData("Already Member");
                    }
                }
                ent.Properties["member"].Add(Obj.Properties["distinguishedName"].Value);
                ent.CommitChanges();


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
            return new ResultData("Done");
        }
    }

}