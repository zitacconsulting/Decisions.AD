using System.Runtime.Serialization;
using System.Collections.Generic;
using System;
using System.DirectoryServices.Protocols;

namespace Zitac.AD.Steps;

[DataContract]
public class Group
{

    [DataMember]
    public string CN { get; set; }

    [DataMember]
    public string Description { get; set; }

    [DataMember]
    public string DistinguishedName { get; set; }

    [DataMember]
    public string ManagedBy { get; set; }

    [DataMember]
    public string Name { get; set; }

    [DataMember]
    public string Email { get; set; }

    [DataMember]
    public string ObjectGUID { get; set; }
    
    [DataMember]
    public string ObjectSID { get; set; }

    [DataMember]
    public string SamAccountName { get; set; }

    [DataMember]
    public DateTime WhenChanged { get; set; }

    [DataMember]
    public DateTime WhenCreated { get; set; }

    [DataMember]
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    public static readonly List<string> GroupAttributes = new List<String> {
    "sAMAccountName",
    "description",
    "mail",
    "managedby",
    "cn",
    "distinguishedname",
    "name",
    "objectguid",
    "objectSid",
    "whenchanged",
    "whencreated"
    };

    public Group()
    {

    }
    public Group(SearchResultEntry entry, List<String> AdditionalAttributes)
    {
        this.SamAccountName = Converters.GetStringProperty(entry, "samaccountname");
        this.Description = Converters.GetStringProperty(entry, "description");
        this.Email = Converters.GetStringProperty(entry, "mail");
        this.ManagedBy = Converters.GetStringProperty(entry, "managedby");
        this.CN = Converters.GetStringProperty(entry, "cn");
        this.DistinguishedName = Converters.GetStringProperty(entry, "distinguishedname");
        this.Name = Converters.GetStringProperty(entry, "name");
        try
        {
            this.ObjectGUID = new Guid((System.Byte[])Converters.GetBinaryProperty(entry, "objectguid")).ToString();
        }
        catch { }
        this.ObjectSID = Converters.GetSIDProperty(entry, "objectSid");
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

    private bool IsEnabled(SearchResultEntry entry)
    {
        var property = entry.Attributes["userAccountControl"];
        if (property != null && property.Count != 0)
        {
            int flags = Int32.Parse((string)property[0]);

            return !Convert.ToBoolean(flags & 0x0002);
        }
        return new bool();
    }

}