using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Flow.CoreSteps;
using DecisionsFramework.Design.Flow.Mapping;
using System.Collections.Generic;

namespace Zitac.AD.Steps;
public class DefaultADProperties : BaseFlowAwareStep
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
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Computer Name"));
                            dataDescriptionList.Add(new DataDescription((DecisionsType) new DecisionsNativeType(typeof (string)), "Additional Attributes", true, true, true));
                            return dataDescriptionList.ToArray();                                              
                        }
            }

            public override OutcomeScenarioData[] OutcomeScenarios {
                get {

                    return new[] {
                    new OutcomeScenarioData("Done", new DataDescription(typeof(Computer), "Result")),
                    new OutcomeScenarioData("Error", new DataDescription(typeof(string), "Error Message")), 
                }; 
                }
            }

}