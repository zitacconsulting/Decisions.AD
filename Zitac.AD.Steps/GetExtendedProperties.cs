using System;
using System.DirectoryServices;

namespace Zitac.AD.Steps;

public class GetExtendedProperties
{
        public Object DynamicallyChoosePropertyGetter(SearchResult entry, string propertyName)
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
    public string GetStringProperty(SearchResult entry, string propertyName)
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
    public DateTime GetDateTimeProperty(SearchResult entry, string propertyName)
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
    public Int64 GetIntProperty(SearchResult entry, string propertyName)
    {
      if (entry != null && !string.IsNullOrEmpty(propertyName))
      {
        ResultPropertyValueCollection property = entry.Properties[propertyName];
        if (property != null && property.Count != 0)
        {
            if(property[0].GetType().ToString() == "System.Int32" || property[0].GetType().ToString() == "System.Int64")
            {
                return (Int64) property[0];
            }
          return new Int64();
        }
        return new Int64();
      }
    return new Int64();
    }
    public System.Byte[] GetBinaryProperty(SearchResult entry, string propertyName)
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
