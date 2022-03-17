using System.Runtime.Serialization;
using System;

namespace Zitac.AD.Steps;

[DataContract]

public class ExtendedAttributes
{
    [DataMember]
    public string Attribute {get; set;}
    
    [DataMember]
    public string StringValue {get; set;}

    [DataMember]
    public Int64 IntegerValue {get; set;}

    [DataMember]
    public DateTime DateValue {get; set;}

    [DataMember]
    public Byte[] BinaryValue {get; set;}

}