using System.DirectoryServices.Protocols;
using System.Runtime.Serialization;
using System.Collections.Generic;
using System;


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

    public static readonly List<string> DCAttributes = new List<String> {
    "name",
    "dNSHostName",
    "siteNames",
};

    public DomainController(SearchResultEntry entry)
    {
 
      this.Name = Converters.GetStringProperty(entry, "name");
      this.dNSHostName = Converters.GetStringProperty(entry, "dNSHostName");
      this.Site = Converters.GetStringProperty(entry, "siteName");
 
    }
  }
