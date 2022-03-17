using ActiveDirectory;
using System.DirectoryServices;
using System.Runtime.Serialization;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using System.Collections.Generic;
using System;
using System.Runtime;

namespace Zitac.AD.Steps;

  [DataContract]
  public class User
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
    public DateTime LastLogonDate { get; set; }

    [DataMember]
    public string Location { get; set; }

    [DataMember]
    public string ManagedBy { get; set; }

    [DataMember]
    public string Name { get; set; }

    [DataMember]
    public string ObjectGUID { get; set; }

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
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    public User(SearchResult entry, string[] AdditionalAttributes)
    {
      this.AccountExpires = this.GetDateTimeProperty(entry, "accountexpires");
      this.CN = this.GetStringProperty(entry, "cn");
      this.Description = this.GetStringProperty(entry, "description");
      this.DistinguishedName = this.GetStringProperty(entry, "distinguishedname");
      this.DNSHostName = this.GetStringProperty(entry, "dnshostname");
      this.LastLogonDate = this.GetDateTimeProperty(entry, "lastLogon");
      this.Location = this.GetStringProperty(entry, "location");
      this.ManagedBy = this.GetStringProperty(entry, "managedby");
      this.Name = this.GetStringProperty(entry, "name");
      this.ObjectGUID = new Guid((System.Byte[])this.GetBinaryProperty(entry, "objectguid")).ToString();
      this.OperatingSystem = this.GetStringProperty(entry, "operatingsystem");
      this.OperatingSystemHotfix = this.GetStringProperty(entry, "operatingsystemhotfix");
      this.OperatingSystemServicePack = this.GetStringProperty(entry, "operatingsystemservicepack");
      this.OperatingSystemVersion = this.GetStringProperty(entry, "operatingsystemversion");
      this.PasswordLastSet = this.GetDateTimeProperty(entry, "pwdlastset");
      this.SamAccountName = this.GetStringProperty(entry, "samaccountname");
      this.WhenChanged = this.GetDateTimeProperty(entry, "whenchanged");
      this.WhenCreated = this.GetDateTimeProperty(entry, "whencreated");
      this.LogonCount = this.GetIntProperty(entry, "logonCount");

      if(AdditionalAttributes != null)
      {
      List<ExtendedAttributes> AttributeResults = new List<ExtendedAttributes>();
      foreach (string Attribute in AdditionalAttributes)
        {
            ExtendedAttributes ToAdd = new ExtendedAttributes();
            ToAdd.Attribute = Attribute;
            ToAdd.StringValue = this.GetStringProperty(entry, Attribute);
            ToAdd.DateValue = this.GetDateTimeProperty(entry, Attribute);
            ToAdd.IntegerValue = this.GetIntProperty(entry, Attribute);
            ToAdd.BinaryValue = this.GetBinaryProperty(entry, Attribute);
            AttributeResults.Add(ToAdd);
        }
      this.AdditionalAttributesResult = AttributeResults.ToArray();
       }
      
    }
    private Object DynamicallyChoosePropertyGetter(SearchResult entry, string propertyName)
    {
      if (entry != null && !string.IsNullOrEmpty(propertyName))
      {
        ResultPropertyValueCollection property = entry.Properties[propertyName];
        if (property != null && property.Count != 0)
        {
          string Type = property[0].GetType().ToString();
          if(Type == "System.Int64" || Type == "System.DateTime")
          {
              return GetDateTimeProperty(entry, propertyName);
          }
          else if(Type == "System.Int32")
          {
              return GetIntProperty(entry, propertyName);
          }
          else if(Type == "System.Byte[]")
          {
              return GetBinaryProperty(entry, propertyName);
          }
          else
          {
              return GetStringProperty(entry, propertyName);
          }
        }
      }
      return null;
    }
    private string GetStringProperty(SearchResult entry, string propertyName)
    {
      if (entry != null && !string.IsNullOrEmpty(propertyName))
      {
        ResultPropertyValueCollection property = entry.Properties[propertyName];
        if (property != null && property.Count != 0)
        {
          return property[0].ToString();
        }
        return (string) null;
      }
    return (string) null;
    }
    private DateTime GetDateTimeProperty(SearchResult entry, string propertyName)
    {
      if (entry != null && !string.IsNullOrEmpty(propertyName))
      {
        ResultPropertyValueCollection property = entry.Properties[propertyName];
        if (property != null && property.Count != 0)
        {
            if(property[0].GetType().ToString() == "System.Int64")
            {
                if(property[0].ToString() == "9223372036854775807")
                {
                    return new DateTime();
                }
                long date = Convert.ToInt64(property[0]);
                System.DateTime FormattedDate = System.DateTime.FromFileTime(date);
                return FormattedDate;
            }
            if(property[0].GetType().ToString() == "System.DateTime")
            {
                return (DateTime) property[0];
            }
          return new DateTime();
        }
        return new DateTime();
      }
    return new DateTime();
    }
    private Int64 GetIntProperty(SearchResult entry, string propertyName)
    {
      if (entry != null && !string.IsNullOrEmpty(propertyName))
      {
        ResultPropertyValueCollection property = entry.Properties[propertyName];
        if (property != null && property.Count != 0)
        {
            if(property[0].GetType().ToString() == "System.Int64")
            {
                return (Int64) property[0];
            }
            if(property[0].GetType().ToString() == "System.Int32"){
              return (Int64) Convert.ToInt64(property[0]);
            }
          return new Int64();
        }
        return new Int64();
      }
    return new Int64();
    }
    private System.Byte[] GetBinaryProperty(SearchResult entry, string propertyName)
    {
      if (entry != null && !string.IsNullOrEmpty(propertyName))
      {
        ResultPropertyValueCollection property = entry.Properties[propertyName];
        if (property != null && property.Count != 0)
        {
            if(property[0].GetType().ToString() == "System.Byte[]")
            {
                return (System.Byte[]) property[0];
            }
          return (System.Byte[]) null;
        }
        return (System.Byte[]) null;
      }
    return (System.Byte[]) null;
    }
  }

