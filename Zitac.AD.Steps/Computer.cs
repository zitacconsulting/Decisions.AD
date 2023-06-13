using System.Runtime.Serialization;
using System.Collections.Generic;
using System;
using System.DirectoryServices.Protocols;

namespace Zitac.AD.Steps;

  [DataContract]
  public class Computer
  {
    [DataMember]
    public DateTime AccountExpires { get; set; }

    [DataMember]
    public string CN { get; set; }

    [DataMember]
    public string Description { get; set; }

    [DataMember]
    public string DistinguishedName { get; set; }

    [DataMember]
    public string DNSHostName { get; set; }

    [DataMember]
    public DateTime LastLogonTimeStamp { get; set; }

    [DataMember]
    public string Location { get; set; }

    [DataMember]
    public string ManagedBy { get; set; }

    [DataMember]
    public string Name { get; set; }

    [DataMember]
    public string ObjectGUID { get; set; }

    [DataMember]
    public string ObjectSID { get; set; }

    [DataMember]
    public string OperatingSystem { get; set; }

    [DataMember]
    public string OperatingSystemHotfix { get; set; }

    [DataMember]
    public string OperatingSystemServicePack { get; set; }

    [DataMember]
    public string OperatingSystemVersion { get; set; }

    [DataMember]
    public DateTime PasswordLastSet { get; set; }

    [DataMember]
    public string SamAccountName { get; set; }


    [DataMember]
    public DateTime WhenChanged { get; set; }

    [DataMember]
    public DateTime WhenCreated { get; set; }
    
    [DataMember]
    public Int64 LogonCount { get; set; }

    [DataMember]
    public bool AccountEnabled { get; set; }

    [DataMember]
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    public static readonly List<string> ComputerAttributes = new List<String> {
    "accountexpires",
    "cn",
    "description",
    "distinguishedname",
    "dnshostname",
    "lastLogonTimestamp",
    "location",
    "managedby",
    "name",
    "objectguid",
    "objectSid",
    "operatingsystem",
    "operatingsystemhotfix",
    "operatingsystemservicepack",
    "operatingsystemversion",
    "pwdlastset",
    "samaccountname",
    "whenchanged",
    "whencreated",
    "logonCount",
    "userAccountControl"
};

    public Computer()
    {
      
    }
    public Computer(SearchResultEntry entry, List<String> AdditionalAttributes)
    {
      this.AccountExpires = Converters.GetDateTimeProperty(entry, "accountexpires");
      this.CN = Converters.GetStringProperty(entry, "cn");
      this.Description = Converters.GetStringProperty(entry, "description");
      this.DistinguishedName = Converters.GetStringProperty(entry, "distinguishedname");
      this.DNSHostName = Converters.GetStringProperty(entry, "dnshostname");
      this.LastLogonTimeStamp = Converters.GetDateTimeProperty(entry, "lastLogonTimestamp");
      this.Location = Converters.GetStringProperty(entry, "location");
      this.ManagedBy = Converters.GetStringProperty(entry, "managedby");
      this.Name = Converters.GetStringProperty(entry, "name");
      this.ObjectGUID = new Guid((System.Byte[])Converters.GetBinaryProperty(entry, "objectguid")).ToString();
      this.ObjectSID = Converters.GetSIDProperty(entry, "objectSid");
      this.OperatingSystem = Converters.GetStringProperty(entry, "operatingsystem");
      this.OperatingSystemHotfix = Converters.GetStringProperty(entry, "operatingsystemhotfix");
      this.OperatingSystemServicePack = Converters.GetStringProperty(entry, "operatingsystemservicepack");
      this.OperatingSystemVersion = Converters.GetStringProperty(entry, "operatingsystemversion");
      this.PasswordLastSet = Converters.GetDateTimeProperty(entry, "pwdlastset");
      this.SamAccountName = Converters.GetStringProperty(entry, "samaccountname");
      this.WhenChanged = Converters.GetDateTimeProperty(entry, "whenchanged");
      this.WhenCreated = Converters.GetDateTimeProperty(entry, "whencreated");
      this.LogonCount = Converters.GetIntProperty(entry, "logonCount");
      this.AccountEnabled = this.IsEnabled(entry);

      if(AdditionalAttributes != null)
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
