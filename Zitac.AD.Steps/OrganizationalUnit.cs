using System.Runtime.Serialization;
using System.Collections.Generic;
using System;
using System.DirectoryServices.Protocols;

namespace Zitac.AD.Steps;

[DataContract]
public class OrganizationalUnit
{

    [DataMember]
    public string Description { get; set; }

    [DataMember]
    public string DistinguishedName { get; set; }

    [DataMember]
    public string ManagedBy { get; set; }

    [DataMember]
    public string Name { get; set; }

    [DataMember]
    public string ObjectGUID { get; set; }
    
    [DataMember]
    public DateTime WhenChanged { get; set; }

    [DataMember]
    public DateTime WhenCreated { get; set; }

    [DataMember]
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    public static readonly List<string> OUAttributes = new List<String> {
    "description",
    "distinguishedname",
    "managedby",
    "name",
    "objectguid",
    "whenchanged",
    "whencreated"
    };

    public OrganizationalUnit()
    {

    }
    public OrganizationalUnit(SearchResultEntry entry, List<String> AdditionalAttributes)
    {
        this.Description = Converters.GetStringProperty(entry, "description");
        this.DistinguishedName = Converters.GetStringProperty(entry, "distinguishedname");
        this.ManagedBy = Converters.GetStringProperty(entry, "managedby");
        this.Name = Converters.GetStringProperty(entry, "name");
        try
        {
            this.ObjectGUID = new Guid((System.Byte[])Converters.GetBinaryProperty(entry, "objectguid")).ToString();
        }
        catch { }
        this.WhenChanged = Converters.GetDateTimeProperty(entry, "whenchanged");
        this.WhenCreated = Converters.GetDateTimeProperty(entry, "whencreated");

        if (AdditionalAttributes != null)
        {
            List<ExtendedAttributes> AttributeResults = new List<ExtendedAttributes>();
            foreach (string Attribute in AdditionalAttributes)
            {
                ExtendedAttributes ToAdd = new ExtendedAttributes();
                ToAdd.Attribute = Attribute;
                ToAdd.StringValue = Converters.GetStringProperty(entry, Attribute);
                ToAdd.DateValue = Converters.GetDateTimeProperty(entry, Attribute);
                ToAdd.IntegerValue = Converters.GetIntProperty(entry, Attribute);
                ToAdd.BinaryValue = Converters.GetBinaryProperty(entry, Attribute);
                AttributeResults.Add(ToAdd);
            }
            this.AdditionalAttributesResult = AttributeResults.ToArray();
        }

    }

}