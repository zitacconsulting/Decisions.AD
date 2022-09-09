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
public class User
{

    [DataMember]
    public string LoginNamePreWin2000 { get; set; }
    
    [DataMember]
    public string LoginNameUPN { get; set; }
    
    [DataMember]
    public string FirstName { get; set; }
    
    [DataMember]
    public string LastName { get; set; }
    
    [DataMember]
    public string Initials { get; set; }

    [DataMember]
    public string DisplayName { get; set; }
    
    [DataMember]
    public string Office { get; set; }
    
    [DataMember]
    public string TelephoneNumber { get; set; }
    
    [DataMember]
    public string Password { get; set; }
    
    [DataMember]
    public string Description { get; set; }
    
    [DataMember]
    public string EmailAddress { get; set; }
    
    [DataMember]
    public string WebPage { get; set; }
    
    [DataMember]
    public DateTime AccountExpires { get; set;}

    [DataMember]
    public string Street { get; set; }
    
    [DataMember]
    public string POBox { get; set; }
    
    [DataMember]
    public string City { get; set; }
    
    [DataMember]
    public string StateProvince { get; set; }
    
    [DataMember]
    public string ZipPostalCode { get; set; }
    
    [DataMember]
    public string CountryRegion { get; set; }
    
    [DataMember]
    public string HomeFolder { get; set; }
    
    [DataMember]
    public string HomeFolderDriveLetter { get; set; }

    [DataMember]
    public string HomePhone { get; set; }

    [DataMember]
    public string MobilePhone { get; set; }

    [DataMember]
    public string Department { get; set; }
    [DataMember]
    public string JobTitle { get; set; }
    
    [DataMember]
    public string Company { get; set; }
    
    [DataMember]
    public string ManagerDN { get; set; }

    [DataMember]
    public string EmployeeID { get; set; }
    
    [DataMember]
    public string EmployeeNumber { get; set; }
    
    [DataMember]
    public string EmployeeType { get; set; }

    [DataMember]
    public string CN { get; set; }

    [DataMember]
    public string DistinguishedName { get; set; }

    [DataMember]
    public DateTime LastLogonDate { get; set; }

    [DataMember]
    public string ObjectGUID { get; set; }

    [DataMember]
    public string ObjectSID { get; set; }

    [DataMember]
    public DateTime PasswordLastSet { get; set; }

    [DataMember]
    public DateTime WhenChanged { get; set; }

    [DataMember]
    public DateTime WhenCreated { get; set; }

    [DataMember]
    public Int64 LogonCount { get; set; }

    [DataMember]
    public Int64 uSNChanged { get; set; }
    
    [DataMember]
    public bool AccountEnabled { get; set; }

    [DataMember]
    public Group[] Groups { get; set; }
    

    [DataMember]
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    private string adString = "";

    public User(SearchResult entry, string[] AdditionalAttributes)
    {

        this.LoginNamePreWin2000 = this.GetStringProperty(entry, "sAMAccountName");
        this.LoginNameUPN = this.GetStringProperty(entry, "userPrincipalName");
        this.FirstName = this.GetStringProperty(entry, "givenName");
        this.LastName = this.GetStringProperty(entry, "sn");
        this.Initials = this.GetStringProperty(entry, "initials");
        this.DisplayName = this.GetStringProperty(entry, "displayName");
        this.Office = this.GetStringProperty(entry, "physicalDeliveryOfficeName");
        this.TelephoneNumber = this.GetStringProperty(entry, "telephoneNumber");
        this.Description = this.GetStringProperty(entry, "description");
        this.EmailAddress = this.GetStringProperty(entry, "mail");
        this.WebPage = this.GetStringProperty(entry, "wWWHomePage");
        this.AccountExpires = this.GetDateTimeProperty(entry, "accountexpires");
        this.Street = this.GetStringProperty(entry, "streetAddress");
        this.POBox = this.GetStringProperty(entry, "postOfficeBox");
        this.City = this.GetStringProperty(entry, "l");
        this.StateProvince = this.GetStringProperty(entry, "st");
        this.ZipPostalCode = this.GetStringProperty(entry, "postalCode");
        this.CountryRegion = this.GetStringProperty(entry, "c");
        this.HomeFolder = this.GetStringProperty(entry, "homeDirectory");
        this.HomeFolderDriveLetter = this.GetStringProperty(entry, "homeDrive");
        this.HomePhone = this.GetStringProperty(entry, "homePhone");
        this.MobilePhone = this.GetStringProperty(entry, "mobile");
        this.Department = this.GetStringProperty(entry, "department");
        this.JobTitle = this.GetStringProperty(entry, "title");
        this.Company = this.GetStringProperty(entry, "company");
        this.ManagerDN = this.GetStringProperty(entry, "manager");
        this.EmployeeID = this.GetStringProperty(entry, "employeeID");
        this.EmployeeNumber = this.GetStringProperty(entry, "employeeNumber");
        this.EmployeeType = this.GetStringProperty(entry, "employeeType");
        this.CN = this.GetStringProperty(entry, "cn");
        this.DistinguishedName = this.GetStringProperty(entry, "distinguishedname");
        this.LastLogonDate = this.GetDateTimeProperty(entry, "lastLogon");
        this.ObjectGUID = new Guid((System.Byte[])this.GetBinaryProperty(entry, "objectguid")).ToString();
        this.ObjectSID = this.GetStringProperty(entry, "objectSid");
        this.PasswordLastSet = this.GetDateTimeProperty(entry, "pwdlastset");
        this.WhenChanged = this.GetDateTimeProperty(entry, "whenchanged");
        this.WhenCreated = this.GetDateTimeProperty(entry, "whencreated");
        this.LogonCount = this.GetIntProperty(entry, "logonCount");
        this.uSNChanged = this.GetIntProperty(entry, "uSNChanged");
        this.AccountEnabled = this.IsEnabled(entry);

        this.Groups = this.GetMembership(entry, "memberOf");
  

        if (AdditionalAttributes != null)
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
                if (Type == "System.Int64" || Type == "System.DateTime")
                {
                    return GetDateTimeProperty(entry, propertyName);
                }
                else if (Type == "System.Int32")
                {
                    return GetIntProperty(entry, propertyName);
                }
                else if (Type == "System.Byte[]")
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
            return (string)null;
        }
        return (string)null;
    }
    private DateTime GetDateTimeProperty(SearchResult entry, string propertyName)
    {
        if (entry != null && !string.IsNullOrEmpty(propertyName))
        {
            ResultPropertyValueCollection property = entry.Properties[propertyName];
            if (property != null && property.Count != 0)
            {
                if (property[0].GetType().ToString() == "System.Int64")
                {
                    if (property[0].ToString() == "9223372036854775807")
                    {
                        return new DateTime();
                    }
                    long date = Convert.ToInt64(property[0]);
                    System.DateTime FormattedDate = System.DateTime.FromFileTime(date);
                    return FormattedDate;
                }
                if (property[0].GetType().ToString() == "System.DateTime")
                {
                    return (DateTime)property[0];
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
                if (property[0].GetType().ToString() == "System.Int64")
                {
                    return (Int64)property[0];
                }
                if (property[0].GetType().ToString() == "System.Int32")
                {
                    return (Int64)Convert.ToInt64(property[0]);
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
                if (property[0].GetType().ToString() == "System.Byte[]")
                {
                    return (System.Byte[])property[0];
                }
                return (System.Byte[])null;
            }
            return (System.Byte[])null;
        }
        return (System.Byte[])null;
    }

    private bool IsEnabled(SearchResult entry) {
            ResultPropertyValueCollection property = entry.Properties["userAccountControl"];
            if (property != null && property.Count != 0)
            {
                int flags = (int)property[0];

                return !Convert.ToBoolean(flags & 0x0002);
            }
            return new bool();
    }

    private Group[] GetMembership(SearchResult entry, string propertyName, bool recursive = false){

        
        
        ResultPropertyValueCollection ValueCollection = entry.Properties[propertyName];
        IEnumerator en = ValueCollection.GetEnumerator();

        ArrayList valuesCollection = null;

        List<Group> GroupList = null;

        while(en.MoveNext())
        {
            if (en.Current != null)
            {
                DirectoryEntry group = new DirectoryEntry(en.Current.ToString());
                DirectoryEntry searchRoot = new DirectoryEntry(baseLdapPath, ADCredentials.ADUsername, ADCredentials.ADPassword);
                DirectorySearcher directorySearcher = new DirectorySearcher(searchRoot);
                directorySearcher.Filter = "(&(objectClass=group)(objectCategory=group)" + Filter + ")";
                //Group GroupResult = new Group(, AdditionalAttributes);

                GroupList.Add(group);
                if (recursive)
                {
                    AttributeValuesMultiString(attributeName, "LDAP://" + en.Current.ToString(), valuesCollection, true);
                }
            
            }
        }
    }

    private List<Group> GetGroupMembership(List<Group> groups, string propertyName, string distinguishedName){
        ResultPropertyValueCollection ValueCollection = entry.Properties[propertyName];
        IEnumerator en = ValueCollection.GetEnumerator();

        ArrayList valuesCollection = null;

        while(en.MoveNext())
        {
            if (en.Current != null)
            {
                if (!valuesCollection.Contains(en.Current.ToString()))
                {
                    valuesCollection.Add(en.Current.ToString());
                    if (recursive)
                    {
                        AttributeValuesMultiString(attributeName, "LDAP://" + en.Current.ToString(), valuesCollection, true);
                    }
                }
            }
        }
    }
}

