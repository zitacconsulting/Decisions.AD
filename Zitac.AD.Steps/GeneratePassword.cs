using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Flow.CoreSteps;
using System.Text;


namespace Zitac.AD.Steps
{
    [AutoRegisterStep("Generate Password", "Integration", "Active Directory", "Zitac")]
    [Writable]
    public class GeneratePassword : BaseFlowAwareStep, ISyncStep, IDataProducer//, INotifyPropertyChanged
    {

        [WritableValue]
        private Int32 requiredLength;

        [WritableValue]
        private bool requireNonLetterOrDigit;

        [WritableValue]
        private bool onlyOneNonLetterOrDigit;

        [WritableValue]
        private bool requireDigit;

        [WritableValue]
        private bool requireLowercase;

        [WritableValue]
        private bool requireUppercase;


        [PropertyClassification(10, "Required Length", new string[] { "Options" })]
        public Int32 RequiredLength
        {
            get { return requiredLength; }
            set
            {
                requiredLength = value;
            }
        }


        [PropertyClassification(4, "Require Non Letter Or Digit", new string[] { "Options" })]
        public bool RequireNonLetterOrDigit
        {
            get { return requireNonLetterOrDigit; }
            set
            {
                requireNonLetterOrDigit = value;
                this.OnPropertyChanged(nameof(RequireNonLetterOrDigit));
                this.OnPropertyChanged("OnlyOneNonLetterOrDigit");

            }
        }

        [BooleanPropertyHidden("RequireNonLetterOrDigit", false)]
        [PropertyClassification(5, "Only one Non Letter Or Digit", new string[] { "Options" })]
        public bool OnlyOneNonLetterOrDigit
        {
            get { return onlyOneNonLetterOrDigit; }
            set
            {
                onlyOneNonLetterOrDigit = value;
            }
        }
    

        [PropertyClassification(1, "Require Digit", new string[] { "Options" })]
        public bool RequireDigit
        {
            get { return requireDigit; }
            set
            {
                requireDigit = value;
            }
        }

        [PropertyClassification(2, "Require Lowercase", new string[] { "Options" })]
        public bool RequireLowercase
        {
            get { return requireLowercase; }
            set
            {
                requireLowercase = value;
            }
        }

        [PropertyClassification(3, "Require Uppercase", new string[] { "Options" })]
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
            List<char> RequiredChars = new List<char>();

            if (RequireNonLetterOrDigit)
                RequiredChars.Add((char)random.Next(33, 48));
            if (RequireDigit)
                RequiredChars.Add((char)random.Next(48, 58));
            if (RequireLowercase)
                RequiredChars.Add((char)random.Next(97, 123));
            if (RequireUppercase)
                RequiredChars.Add((char)random.Next(65, 91));
            


            while (password.Length < (RequiredLength))
            {
                if ((RequiredLength - password.Length) <= RequiredChars.Count && RequiredChars.Count != 0 ) {
                    var ToAdd = random.Next(0,(RequiredChars.Count));
                    password.Append(RequiredChars.ElementAt(ToAdd));
                    RequiredChars.RemoveAt(ToAdd);
                }
                else if ((random.Next(2) == 1) && RequiredChars.Count != 0 ){
                    var ToAdd = random.Next(0,(RequiredChars.Count));
                    password.Append(RequiredChars.ElementAt(ToAdd));
                    RequiredChars.RemoveAt(ToAdd);
                }
                else {
                    int CharRandom;
                    if (OnlyOneNonLetterOrDigit){
                        CharRandom = random.Next(0,3);
                    }
                    else {
                        CharRandom = random.Next(0,4);
                    }

                    switch(CharRandom) {
                        case 0:
                            // Digit
                            password.Append((char)random.Next(48, 58));
                            break;
                        case 1:
                            // Lowercase
                            password.Append((char)random.Next(97, 123));
                            break;
                        case 2:
                            // Uppercase
                            password.Append((char)random.Next(65, 91));
                            break;
                        case 3:
                            // NonLetterOrDigit
                            password.Append((char)random.Next(33, 47));
                            break;                                                    
                    }
                }

            }

            return new ResultData("Done", (IDictionary<string, object>)new Dictionary<string, object>() { { "Password", (string)password.ToString() } });
        }

    }

}