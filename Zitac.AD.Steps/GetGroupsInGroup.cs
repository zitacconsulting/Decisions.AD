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
    [AutoRegisterStep("Get Groups In Group", "Integration", "Active Directory", "Zitac", "Group")]
    [Writable]
    public class GetGroupsInGroup : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer //, INotifyPropertyChanged
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

                outcomeScenarioDataList.Add(new OutcomeScenarioData("Done", new DataDescription(typeof(Group), "Groups", true)));
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

                List<Group> GroupList = new List<Group>();

                //string PrimaryGroupDn = GetUserPrimaryGroup((DirectoryEntry) one.GetDirectoryEntry());
                GroupList = GetGroupMembers(GroupList, one.Properties["distinguishedname"][0].ToString(), directorySearcher, AdditionalAttributes, RecursiveSearch);

                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                dictionary.Add("Groups", (object)GroupList.ToArray());
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

        private static List<Group> GetGroupMembers(List<Group> GroupList, string DN, DirectorySearcher directorySearcher, string[] AdditionalAttributes, Boolean RecursiveSearch)
        {

            directorySearcher.Filter = "(&(objectClass=group)(objectCategory=group)(memberOf=" + DN + "))";
            SearchResultCollection GroupGroupMembers = directorySearcher.FindAll();

            if (GroupGroupMembers != null && GroupGroupMembers.Count != 0)
            {
                foreach (SearchResult Current in GroupGroupMembers)
                {
                    if (!GroupList.Any(i => i.DistinguishedName == Current.Properties["distinguishedname"][0].ToString()))
                    {
                        GroupList.Add(new Group(Current, AdditionalAttributes));
                        if (RecursiveSearch)
                        {
                            GroupList = GetGroupMembers(GroupList, Current.Properties["distinguishedname"][0].ToString(), directorySearcher, AdditionalAttributes, RecursiveSearch);
                        }
                    }
                }
            }
            return GroupList;
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
}