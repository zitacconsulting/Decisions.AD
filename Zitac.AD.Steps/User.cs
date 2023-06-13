
using System.Runtime.Serialization;
using System.Collections.Generic;
using System;
using System.DirectoryServices.Protocols;

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
    public string Description { get; set; }

    [DataMember]
    public string EmailAddress { get; set; }

    [DataMember]
    public string WebPage { get; set; }

    [DataMember]
    public DateTime AccountExpires { get; set; }

    [DataMember]
    public bool PasswordNeverExpires { get; set; }

    [DataMember]
    public DateTime? PasswordExpiration { get; set; }

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
    public DateTime PwdLastSet { get; set; }

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
    public ExtendedAttributes[] AdditionalAttributesResult { get; set; }

    public static readonly List<string> UserAttributes = new List<String> {
    "sAMAccountName",
    "userPrincipalName",
    "givenName",
    "sn",
    "initials",
    "displayName",
    "physicalDeliveryOfficeName",
    "telephoneNumber",
    "description",
    "mail",
    "wWWHomePage",
    "accountexpires",
    "streetAddress",
    "postOfficeBox",
    "l",
    "st",
    "postalCode",
    "c",
    "homeDirectory",
    "homeDrive",
    "homePhone",
    "mobile",
    "department",
    "title",
    "company",
    "manager",
    "employeeID",
    "employeeNumber",
    "employeeType",
    "cn",
    "distinguishedname",
    "lastLogon",
    "objectguid",
    "objectSid",
    "pwdlastset",
    "whenchanged",
    "whencreated",
    "logonCount",
    "uSNChanged",
    "userAccountControl"
};

    public User()
    {
    }
    public User(SearchResultEntry entry, List<String> AdditionalAttributes, Int32 PwdExpDays)
    {

        this.LoginNamePreWin2000 = Converters.GetStringProperty(entry, "sAMAccountName");
        this.LoginNameUPN = Converters.GetStringProperty(entry, "userPrincipalName");
        this.FirstName = Converters.GetStringProperty(entry, "givenName");
        this.LastName = Converters.GetStringProperty(entry, "sn");
        this.Initials = Converters.GetStringProperty(entry, "initials");
        this.DisplayName = Converters.GetStringProperty(entry, "displayName");
        this.Office = Converters.GetStringProperty(entry, "physicalDeliveryOfficeName");
        this.TelephoneNumber = Converters.GetStringProperty(entry, "telephoneNumber");
        this.Description = Converters.GetStringProperty(entry, "description");
        this.EmailAddress = Converters.GetStringProperty(entry, "mail");
        this.WebPage = Converters.GetStringProperty(entry, "wWWHomePage");
        this.AccountExpires = Converters.GetDateTimeProperty(entry, "accountexpires");
        this.PasswordNeverExpires = this.GetNeverExpires(entry);
        this.PasswordExpiration = this.GetPasswordExpiration(entry, PwdExpDays);
        this.Street = Converters.GetStringProperty(entry, "streetAddress");
        this.POBox = Converters.GetStringProperty(entry, "postOfficeBox");
        this.City = Converters.GetStringProperty(entry, "l");
        this.StateProvince = Converters.GetStringProperty(entry, "st");
        this.ZipPostalCode = Converters.GetStringProperty(entry, "postalCode");
        this.CountryRegion = Converters.GetStringProperty(entry, "c");
        this.HomeFolder = Converters.GetStringProperty(entry, "homeDirectory");
        this.HomeFolderDriveLetter = Converters.GetStringProperty(entry, "homeDrive");
        this.HomePhone = Converters.GetStringProperty(entry, "homePhone");
        this.MobilePhone = Converters.GetStringProperty(entry, "mobile");
        this.Department = Converters.GetStringProperty(entry, "department");
        this.JobTitle = Converters.GetStringProperty(entry, "title");
        this.Company = Converters.GetStringProperty(entry, "company");
        this.ManagerDN = Converters.GetStringProperty(entry, "manager");
        this.EmployeeID = Converters.GetStringProperty(entry, "employeeID");
        this.EmployeeNumber = Converters.GetStringProperty(entry, "employeeNumber");
        this.EmployeeType = Converters.GetStringProperty(entry, "employeeType");
        this.CN = Converters.GetStringProperty(entry, "cn");
        this.DistinguishedName = Converters.GetStringProperty(entry, "distinguishedname");
        this.LastLogonDate = Converters.GetDateTimeProperty(entry, "lastLogon");
        this.ObjectGUID = new Guid((System.Byte[])Converters.GetBinaryProperty(entry, "objectguid")).ToString();
        this.ObjectSID = Converters.GetSIDProperty(entry, "objectSid");
        this.PwdLastSet = Converters.GetDateTimeProperty(entry, "pwdlastset");
        this.WhenChanged = Converters.GetDateTimeProperty(entry, "whenchanged");
        this.WhenCreated = Converters.GetDateTimeProperty(entry, "whencreated");
        this.LogonCount = Converters.GetIntProperty(entry, "logonCount");
        this.uSNChanged = Converters.GetIntProperty(entry, "uSNChanged");
        this.AccountEnabled = this.IsEnabled(entry);

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
private bool GetNeverExpires(SearchResultEntry entry)
{
    var property = entry.Attributes["userAccountControl"];
    if (property != null && property.Count != 0)
    {
        string hej = (string)property[0];
        int flags = Int32.Parse((string)property[0]);

        return Convert.ToBoolean(flags & 0x10000);
    }
    return new bool();
}
private DateTime? GetPasswordExpiration(SearchResultEntry entry, Int32 PwdExpDays)
{
    if (GetNeverExpires(entry))
    {
        return null;
    }
    else
    {
        DateTime ExpirationDate = Converters.GetDateTimeProperty(entry, "pwdlastset").AddDays(PwdExpDays);
        return ExpirationDate;
    }
}
}