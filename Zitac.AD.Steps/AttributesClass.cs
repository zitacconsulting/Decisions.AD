namespace Zitac.AD.Steps
{
    public class AttributeValues
    {
        public string Parameter { get; set; }
        public string Attribute { get; set; }

        public AttributeValues(string parameter, string attribute)
        {
            Parameter = parameter;
            Attribute = attribute;
        }
    }
}