
using System.Runtime.Serialization;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Properties;
using System.ComponentModel;

namespace Zitac.AD.Steps;

[Writable]
[DataContract]
public class SearchParameters
{

    private string fieldName;
    [WritableValue]
    private string alias;

    [WritableValue]
    [DataMember]
    [PropertyClassification("Query Match Type", 2)]
    [SelectStringEditor("MatchValues")]
    public string MatchCriteria { get; set; }

    [PropertyHidden]
    public string[] MatchValues
    {
      get
      {
          return new[] {
                  "Equals", "Contains", "DoesNotContain", "DoesNotEqual", "GreaterThanOrEqualTo", "LessThanOrEqualTo", "GreaterThan", "LessThan", "Exists", "DoesNotExist", "StartsWith", "DoesNotStartWith", "EndsWith", "DoesNotEndWith"
          };
      }
      set { return; }
    }

    [WritableValue]
    [DataMember]
    [PropertyClassification("Field Data Type", 2)]
    [SelectStringEditor("DataTypes")]
    public string DataType { get; set; }

    [PropertyHidden]
    public string[] DataTypes
    {
      get
      {
          return new[] {
                  "String", "Date", "Int32", "Int64"
          };
      }
      set { return; }
    }

    [WritableValue]
    [DataMember]
    [PropertyClassification("Field Name", 1)]
    public string FieldName
    {
      get => this.fieldName;
      set
      {
        this.fieldName = value;
      }
    }

    
    [DataMember]
    [PropertyClassification("Input Data Alias (Optional)", 3)]
    public string Alias
    {
      get
      {
        if (string.IsNullOrEmpty(this.alias))
          this.alias = this.FieldName;
        return this.alias;
      }
      set => this.alias = value;
    }

    [PropertyHidden]
    public string SearchValue { get; set; }

}
