using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.DirectoryServices.Protocols;

namespace Zitac.AD.Steps
{
    public class Converters
    {
        // Convert a binary SID to a string
        public static string ConvertSidToString(IEnumerable<byte> byteCollection)
        {
            // sid[0] is the Revision, we allow only version 1, because it's the
            // only version that exists right now.
            if (byteCollection.ElementAt(0) != 1)
            {
                throw new ArgumentOutOfRangeException("SID (bytes[0]) revision must be 1");
            }

            var stringSidBuilder = new StringBuilder("S-1-");

            // The next byte specifies the numbers of sub authorities
            // (number of dashes minus two), should be 5 or less, but not enforcing that
            var subAuthorityCount = byteCollection.ElementAt(1);

            // IdentifierAuthority (6 bytes starting from the second) (big endian)
            long identifierAuthority = 0;

            var offset = 2;
            var size = 6;
            for (var i = 0; i < size; i++)
            {
                identifierAuthority |= (long)byteCollection.ElementAt(offset + i) << (8 * (size - 1 - i));
            }

            stringSidBuilder.Append(identifierAuthority.ToString());

            // Iterate all the SubAuthority (little-endian)
            offset = 8;
            size = 4; // 32-bits (4 bytes) for each SubAuthority
            var j = 0;
            for (var i = 0; i < subAuthorityCount; i++)
            {
                long subAuthority = 0;

                for (j = 0; j < size; j++)
                {
                    // the below "Or" is a logical Or not a boolean operator
                    subAuthority |= (long)byteCollection.ElementAt(offset + j) << (8 * j);
                }
                stringSidBuilder.Append("-").Append(subAuthority);
                offset += size;
            }

            return stringSidBuilder.ToString();
        }

        public static string GetStringProperty(SearchResultEntry entry, string propertyName)
        {
            if (entry != null && !string.IsNullOrEmpty(propertyName))
            {
                var property = entry.Attributes[propertyName];
                if (property != null && property.Count != 0)
                {
                    return property[0].ToString();
                }
                return (string)null;
            }
            return (string)null;
        }

        public static string[] GetStringListProperty(SearchResultEntry entry, string propertyName)
        {
        DirectoryAttribute memberOfAttribute = entry.Attributes[propertyName];
        if (memberOfAttribute != null)
        {
            string[] ListResults = new string[memberOfAttribute.Count];
            for (int i = 0; i < memberOfAttribute.Count; i++)
            {
                ListResults[i] = memberOfAttribute[i].ToString();
            }
            return ListResults;
        }
        else {
            string[] ListResults = new string[0];
            return ListResults;
        }
        }
        public static DateTime GetDateTimeProperty(SearchResultEntry entry, string propertyName)
        {
            if (entry != null && !string.IsNullOrEmpty(propertyName))
            {
                var property = entry.Attributes[propertyName];
                if (property != null && property[0] != null)
                {
                    string value = (string)property[0];
                    if (value == "9223372036854775807")
                    {
                        return new DateTime();
                    }
                    else if (long.TryParse(value, out long fileTime))
                    {
                        DateTime ParsedDate = DateTime.FromFileTime(fileTime);
                        return ParsedDate;
                    }
                    else if (DateTime.TryParseExact(value, "yyyyMMddHHmmss.0Z", null, System.Globalization.DateTimeStyles.AssumeLocal, out DateTime ParsedDate))
                    {
                        return ParsedDate;
                    }
                    else
                    {
                        return new DateTime();
                    }
                }
                return new DateTime();
            }
            return new DateTime();
        }
        public static Int64 GetIntProperty(SearchResultEntry entry, string propertyName)
        {
            if (entry != null && !string.IsNullOrEmpty(propertyName))
            {
                var property = entry.Attributes[propertyName];
                if (property != null)
                {
                    try
                    {
                        return (Int64)Convert.ToInt64(property[0]);
                    }
                    catch
                    {
                        return new Int64();
                    }
                }
                return new Int64();
            }
            return new Int64();
        }
        public static System.Byte[] GetBinaryProperty(SearchResultEntry entry, string propertyName)
        {
            if (entry != null && !string.IsNullOrEmpty(propertyName))
            {
                var property = entry.Attributes[propertyName];
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

        public static String GetSIDProperty(SearchResultEntry entry, string propertyName)
        {
            if (entry != null && !string.IsNullOrEmpty(propertyName))
            {
                try
                {
                    var property = entry.Attributes[propertyName];
                    if (property[0] != null)
                    {
                        return Converters.ConvertSidToString((System.Byte[])(Array)property[0]);
                    }
                    return (string)null;


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    return (string)null;
                }

            }
            return (string)null;
        }
    }


}