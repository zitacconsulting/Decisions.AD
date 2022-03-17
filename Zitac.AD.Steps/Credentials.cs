using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.Service.Debugging.DebugData;
using System;
using System.Collections.Generic;


namespace Zitac.AD.Steps;

        public class Credentials : IDebuggerJsonProvider
    {
        [WritableValue]
        public string ADUsername { get; set; }

        [WritableValue]
        [PasswordText]
        public string ADPassword { get; set; }

        public object GetJsonDebugView()
        {
                return new
                {
                   ADUsername = this.ADUsername,
                   ADPassword = "********"
                };
        }
    }
