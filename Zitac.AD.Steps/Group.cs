using ActiveDirectory;
using System.DirectoryServices;
using System.Runtime.Serialization;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using System.Collections.Generic;
using System.Collections;
using System;
using System.Runtime;

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
    public string SamAccountName { get; set; }

    [DataMember]
    public DateTime WhenChanged { get; set; }

    [DataMember]
    public DateTime WhenCreated { get; set; }

    [DataMember]
    public Group[] MemberOf { get; set; }
    
    [DataMember]
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    public Group(SearchResult entry, string[] AdditionalAttributes, string ADServer, string ADUsername, string ADPassword, bool recursive, List<String> PreviouslyProcessedDN)
    {
      this.SamAccountName = this.GetStringProperty(entry, "samaccountname");
      this.Description = this.GetStringProperty(entry, "description");
      this.Email = this.GetStringProperty(entry, "mail");
      this.ManagedBy = this.GetStringProperty(entry, "managedby");
      this.CN = this.GetStringProperty(entry, "cn");
      this.DistinguishedName = this.GetStringProperty(entry, "distinguishedname");
      this.Name = this.GetStringProperty(entry, "name");
      this.ObjectGUID = new Guid((System.Byte[])this.GetBinaryProperty(entry, "objectguid")).ToString();
      this.WhenChanged = this.GetDateTimeProperty(entry, "whenchanged");
      this.WhenCreated = this.GetDateTimeProperty(entry, "whencreated");

      GroupHelper gr = new GroupHelper(); 
      this.MemberOf = gr.GetMembership(entry, "memberOf", recursive, ADServer, ADUsername, ADPassword, PreviouslyProcessedDN);

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
public class GroupHelper {
    public Group[] GetMembership(SearchResult entry, string propertyName, bool recursive, string ADServer, string ADUsername, string ADPassword, List<String> PreviouslyProcessedDN)
    {
        ResultPropertyValueCollection ValueCollection = entry.Properties[propertyName];
        IEnumerator en = ValueCollection.GetEnumerator();

        List<Group> GroupList = new List<Group>();

        while (en.MoveNext())
        {
            if (en.Current != null & !PreviouslyProcessedDN.Contains(en.Current.ToString().ToUpper()))
            {
                PreviouslyProcessedDN.Add(en.Current.ToString().ToUpper());
                GroupList = GetGroupMembership(GroupList, propertyName, en.Current.ToString(), recursive, ADServer, ADUsername, ADPassword, PreviouslyProcessedDN);
            }
        }
        return GroupList.ToArray();
    }

    public List<Group> GetGroupMembership(List<Group> groups, string propertyName, string distinguishedName, bool recursive, string ADServer, string ADUsername, string ADPassword, List<String> PreviouslyProcessedDN)
    {
        DirectoryEntry searchRoot = new DirectoryEntry("LDAP://" + ADServer, ADUsername, ADPassword);
        DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
        directorySearcher.Filter = "(&(objectClass=group)(objectCategory=group)(distinguishedname=" + distinguishedName + "))";
        SearchResult one = directorySearcher.FindOne();
        if (searchRoot != null)
        {
            searchRoot.Close();
            searchRoot.Dispose();
        }
        directorySearcher.Dispose();
        Group group = new Group(one, null, ADServer, ADUsername, ADPassword, recursive, null);
        if (!groups.Contains(group))
        {
            groups.Add(group);
            if (recursive)
            {
              foreach(Group SubGroup in group.MemberOf )
              {
                  groups = GetGroupMembership(groups, propertyName, SubGroup.DistinguishedName, recursive, ADServer, ADUsername, ADPassword, PreviouslyProcessedDN);
              }
                
            }
        }
        return groups;
    }
}