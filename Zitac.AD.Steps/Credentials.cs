
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Service.Debugging.DebugData;
using System.Runtime.Serialization;


namespace Zitac.AD.Steps;
    [Writable]
    [DataContract]
        public class Credentials : IDebuggerJsonProvider
    {
        [WritableValue]
        public string Username { get; set; }
        [WritableValue]
        public string Domain { get; set; }

        [WritableValue]
        [PasswordText]
        public string Password { get; set; }

        public object GetJsonDebugView()
        {
                return new
                {
                   Username = Username,
                   Domain = Domain,
                   Password = "********"
                };
        }
    }
