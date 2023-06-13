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
    [AutoRegisterStep("Remove From Group", "Integration", "Active Directory", "Zitac", "Group")]
    [Writable]
    public class RemoveFromGroup : BaseFlowAwareStep, ISyncStep, IDataConsumer, IDataProducer, IDefaultInputMappingStep //, INotifyPropertyChanged
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
            get
            {

                List<DataDescription> dataDescriptionList = new List<DataDescription>();
                if (!IntegratedAuthentication)
                {
                    dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(Credentials)), "Credentials"));
                }

                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(string)), "AD Server"));
                dataDescriptionList.Add(new DataDescription((DecisionsType)new DecisionsNativeType(typeof(int?)), "Port", false, true, false));
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
                    new OutcomeScenarioData("Not Member"),
                    new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")),
                };
            }
        }

        public ResultData Run(StepStartData data)
        {
            Dictionary<string, object> resultData = new Dictionary<string, object>();
            string ADServer = data.Data["AD Server"] as string;
            int? Port = (int?)data.Data["Port"];
            string Group = data.Data["Group Name or DN"] as string;
            string Object = data.Data["Object Account Name or DN"] as string;

            string Filter = string.Empty;

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
            Filter = "(&(objectCategory=group)(objectClass=group)(|(sAMAccountName=" + Group + ")(distinguishedName=" + Group + ")))";
            try
            {
                string FoundGroup = string.Empty;
                string FoundObject = string.Empty;
                IntegrationOptions Options = new IntegrationOptions(ADServer, Port, ADCredentials, UseSSL, IgnoreInvalidCert, IntegratedAuthentication);


                LdapConnection connection = LDAPHelper.GenerateLDAPConnection(Options);
                string BaseDN = LDAPHelper.GetBaseDN(connection);
                List<SearchResultEntry> GroupResults = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "distinguishedname", "objectSid" }).ToList();

                if (GroupResults != null && GroupResults.Count != 0)
                {
                    FoundGroup = Converters.GetStringProperty(GroupResults[0], "distinguishedname");
                }
                else
                {
                    throw new Exception(string.Format("Unable to find group with name or DN: '{0}' in the AD", (object)Group));
                }

                Filter = "(|(sAMAccountName=" + Object + ")(distinguishedName=" + Object + "))";
                List<SearchResultEntry> ObjectResults = LDAPHelper.GetPagedLDAPResults(connection, BaseDN, SearchScope.Subtree, Filter, new List<string> { "distinguishedname", "primaryGroupID", "memberof" }).ToList();

                if (ObjectResults != null && ObjectResults.Count != 0)
                {
                    FoundObject = Converters.GetStringProperty(ObjectResults[0], "distinguishedname");
                }
                else
                {
                    throw new Exception(string.Format("Unable to find Object with name or DN: '{0}' in the AD", (object)Object));
                }

                foreach (string g in Converters.GetStringListProperty(ObjectResults[0], "memberof"))
                {
                    Console.WriteLine(g + " - " + FoundGroup);
                    if (g.Equals(FoundGroup))
                    {
                        DirectoryAttributeModification attributeModification = new DirectoryAttributeModification();
                        attributeModification.Name = "member";
                        attributeModification.Operation = DirectoryAttributeOperation.Delete;
                        attributeModification.Add(FoundObject);
                        ModifyRequest modifyRequest = new ModifyRequest(FoundGroup, new DirectoryAttributeModification[1]
                        {
                    attributeModification
                        });
                        ModifyResponse response = (ModifyResponse)connection.SendRequest((DirectoryRequest)modifyRequest);

                        if (response.ResultCode != ResultCode.Success)
                        {
                            throw new Exception("Failed to remove user from group. Result code: " + response.ResultCode + ". Error: " + response.ErrorMessage);
                        }
                        return new ResultData("Done");
                    }
                }
                return new ResultData("Not Member");
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