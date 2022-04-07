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
using System.Text;


namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Generate Password", "Integration", "Active Directory", "Zitac")]
    [Writable]
    public class GeneratePassword : BaseFlowAwareStep, ISyncStep, IDataProducer //, INotifyPropertyChanged
    {

        [WritableValue]
        private Int32 requiredLength;

        [WritableValue]
        private bool requireNonLetterOrDigit;

        [WritableValue]
        private bool requireDigit;

        [WritableValue]
        private bool requireLowercase;

        [WritableValue]
        private bool requireUppercase;


        [PropertyClassification(10, "RequiredLength", new string[] { "Options" })]
        public Int32 RequiredLength
        {
            get { return requiredLength; }
            set
            {
                requiredLength = value;
            }
        }


        [PropertyClassification(10, "Require Non Letter Or Digit", new string[] { "Options" })]
        public bool RequireNonLetterOrDigit
        {
            get { return requireNonLetterOrDigit; }
            set
            {
                requireNonLetterOrDigit = value;
            }
        }

        [PropertyClassification(10, "Require Digit", new string[] { "Options" })]
        public bool RequireDigit
        {
            get { return requireDigit; }
            set
            {
                requireDigit = value;
            }
        }

        [PropertyClassification(10, "Require Lowercase", new string[] { "Options" })]
        public bool RequireLowercase
        {
            get { return requireLowercase; }
            set
            {
                requireLowercase = value;
            }
        }

        [PropertyClassification(10, "Require Uppercase", new string[] { "Options" })]
        public bool RequireUppercase
        {
            get { return requireUppercase; }
            set
            {
                requireUppercase = value;
            }
        }


        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {

                return new[] {
                    new OutcomeScenarioData("Done", new DataDescription(typeof(string), "Password"))
                 };
            }
        }

        public ResultData Run(StepStartData data)
        {

            StringBuilder password = new StringBuilder();
            Random random = new Random();

            while (password.Length < RequiredLength)
            {
                char c = (char)random.Next(32, 126);

                password.Append(c);

                if (char.IsDigit(c))
                    RequireDigit = false;
                else if (char.IsLower(c))
                    RequireLowercase = false;
                else if (char.IsUpper(c))
                    requireUppercase = false;
                else if (!char.IsLetterOrDigit(c))
                    RequireNonLetterOrDigit = false;
            }

            if (RequireNonLetterOrDigit)
                password.Append((char)random.Next(33, 48));
            if (RequireDigit)
                password.Append((char)random.Next(48, 58));
            if (RequireLowercase)
                password.Append((char)random.Next(97, 123));
            if (requireUppercase)
                password.Append((char)random.Next(65, 91));

            return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "Password", (string)password.ToString() } });
        }

    }

}