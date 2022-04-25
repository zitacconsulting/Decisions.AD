using ActiveDirectory;
using System.DirectoryServices;
using System.Runtime.Serialization;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using System.Collections.Generic;
using System;
using System.Runtime;

namespace Zitac.AD.Steps;

  [DataContract]
  public class DomainController
  {


    [DataMember]
    public string Name { get; set; }

    [DataMember]
    public string dNSHostName { get; set; }

    [DataMember]
    public string Site { get; set; }

    public DomainController(SearchResult entry)
    {
 
      this.Name = this.GetStringProperty(entry, "name");
      this.dNSHostName = this.GetStringProperty(entry, "dNSHostName");
      this.Site = this.GetStringProperty(entry, "siteName");
 
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

