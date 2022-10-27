using ActiveDirectory;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Service.Debugging.DebugData;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using System;
using System.Collections.Generic;
using System.Collections;
using System.DirectoryServices;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Security.Permissions;
using System.Security.Principal;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.ComponentModel;

namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Get Computers In Group", "Integration", "Active Directory", "Zitac", "Group")]
    [Writable]
    public class GetComputersInGroup : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer //, INotifyPropertyChanged
    {

        [WritableValue]
        private bool integratedAuthentication;

        [WritableValue]
        private bool showOutcomeforNoResults;

        [WritableValue]
        private bool recursiveSearch;

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

        [PropertyClassification(new string[] { "Recursive Search" })]
        public bool RecursiveSearch
        {
            get { return recursiveSearch; }
            set
            {
                recursiveSearch = value;
            }
        }

        [PropertyClassification(1, "Show Outcome for No Results", new string[] { "Outcomes" })]
        public bool ShowOutcomeforNoResults
        {
            get { return showOutcomeforNoResults; }
            set
            {
                showOutcomeforNoResults = value;
                this.OnPropertyChanged("OutcomeScenarios");
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
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Group name or DN"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "Additional Attributes", true, true, true));
                return dataDescriptionList.ToArray();
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                List<OutcomeScenarioData> outcomeScenarioDataList = new List<OutcomeScenarioData>();

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(Computer), "Computers", true)));
                if (ShowOutcomeforNoResults)
                {
                    outcomeScenarioDataList.Add(new OutcomeScenarioData("No Results"));
                }
                outcomeScenarioDataList.Add(new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")));
                return outcomeScenarioDataList.ToArray();
            }
        }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            string GroupName = data.Data["Group name or DN"] as string;
            string[] AdditionalAttributes = data.Data["Additional Attributes"] as string[];

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
                directorySearcher.Filter = "(&(objectClass=group)(objectCategory=group)(|(sAMAccountName=" + GroupName + ")(distinguishedname=" + GroupName + ")))";

                SearchResult one = directorySearcher.FindOne();

                if (searchRoot != null)
                {
                    searchRoot.Close();
                    searchRoot.Dispose();
                }
                directorySearcher.Dispose();

                if (one == null)
                {
                    if (ShowOutcomeforNoResults)
                    {
                        return new ResultData("No Results");
                    }
                    throw new Exception(string.Format("Unable to find group with name: '{0}' in the AD", (object)GroupName));

                }
                ComputerAndGroupList ComputerAndGroupList = new ComputerAndGroupList();
                ComputerAndGroupList.GroupList = new List<Group>();
                ComputerAndGroupList.ComputerList = new List<Computer>();
                
                string SID = new SecurityIdentifier((byte[])one.Properties["objectSid"][0], 0).ToString();
                string RID = SID.Substring(SID.LastIndexOf("-", StringComparison.Ordinal) + 1);

                ComputerAndGroupList = GetComputerGroupMembers(ComputerAndGroupList, one.Properties["distinguishedname"][0].ToString(), RID, directorySearcher, AdditionalAttributes, RecursiveSearch);

                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("Computers", (object)ComputerAndGroupList.ComputerList.ToArray());
                return new ResultData("Done", (IDictionary<string, object>)dictionary);

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

        private static ComputerAndGroupList GetComputerGroupMembers(ComputerAndGroupList ComputerAndGroupList, string DN, string RID, DirectorySearcher directorySearcher, string[] AdditionalAttributes, Boolean RecursiveSearch)
        {

            directorySearcher.Filter = "(&(objectClass=group)(objectCategory=group)(memberOf=" + DN + "))";
            SearchResultCollection GroupGroupMembers = directorySearcher.FindAll();

            if (GroupGroupMembers != null && GroupGroupMembers.Count != 0)
            {
                foreach (SearchResult Current in GroupGroupMembers)
                {
                    if (!ComputerAndGroupList.GroupList.Any(i => i.DistinguishedName == Current.Properties["distinguishedname"][0].ToString()))
                    {
                        ComputerAndGroupList.GroupList.Add(new Group(Current, null));
                        if (RecursiveSearch)
                        {
                            string SID = new SecurityIdentifier((byte[])Current.Properties["objectSid"][0], 0).ToString();
                            string CurrentRID = SID.Substring(SID.LastIndexOf("-", StringComparison.Ordinal) + 1);
                            ComputerAndGroupList = GetComputerGroupMembers(ComputerAndGroupList, Current.Properties["distinguishedname"][0].ToString(),CurrentRID, directorySearcher, AdditionalAttributes, RecursiveSearch);
                        }
                    }
                }
            }
            
            directorySearcher.Filter = "(&(objectClass=computer)(|(memberOf=" + DN + ")(primaryGroupId=" + RID + ")))";
            SearchResultCollection ComputerGroupMembers = directorySearcher.FindAll();

            if (ComputerGroupMembers != null && ComputerGroupMembers.Count != 0)
            {
                foreach (SearchResult Current in ComputerGroupMembers)
                {
                    if (!ComputerAndGroupList.ComputerList.Any(i => i.DistinguishedName == Current.Properties["distinguishedname"][0].ToString()))
                    {
                        ComputerAndGroupList.ComputerList.Add(new Computer(Current, AdditionalAttributes));
                    }
                }
            }

            return ComputerAndGroupList;
        }

        private static string GetUserPrimaryGroup(DirectoryEntry de)
        {
            de.RefreshCache(new[] { "primaryGroupID", "objectSid" });

            //Get the user's SID as a string
            var sid = new SecurityIdentifier((byte[])de.Properties["objectSid"].Value, 0).ToString();

            //Replace the RID portion of the user's SID with the primaryGroupId
            //so we're left with the group's SID
            sid = sid.Remove(sid.LastIndexOf("-", StringComparison.Ordinal) + 1);
            sid = sid + de.Properties["primaryGroupId"].Value;

            //Find the group by its SID
            var group = new DirectoryEntry($"LDAP://<SID={sid}>");
            group.RefreshCache(new[] { "distinguishedname" });

            return group.Properties["distinguishedname"].Value as string;
        }
    }

    public class ComputerAndGroupList {
        public List<Computer> ComputerList;
        public List<Group> GroupList;
    }
}