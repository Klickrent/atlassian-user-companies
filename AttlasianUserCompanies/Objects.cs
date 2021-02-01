using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace AtlassianUserCompanies
{
    public class User
    {
        public string name { get; set; }
        public string key { get; set; }
        public string emailAddress { get; set; }
        public string displayName { get; set; }
        public bool active { get; set; }
        public string timeZone { get; set; }
        public bool jira { get; set; }
        public bool confluence { get; set; }
        public Company company { get; set; }
        public int countLicenceGroups { get; set; }
    }

    public class Users
    {
        public int size { get; set; }
        public List<User> items { get; set; }

        [JsonPropertyName("max-results")]
        public int maxResults { get; set; }

        [JsonPropertyName("start-index")]
        public int startIndex { get; set; }

        [JsonPropertyName("end-index")]
        public int endIndex { get; set; }
    }

    public class JiraGroupRootObject
    {
        public string name { get; set; }
        public string self { get; set; }
        public Users users { get; set; }
        public string expand { get; set; }
    }

    public class UserCompany
    {
        public int valueAssociationId { get; set; }

        [JsonPropertyName("inputOptionId")]
        public int CompanyID { get; set; }

        [JsonPropertyName("inputOptionName")]
        public string CompanyName { get; set; }

    }

    public class UserCompanyRootObject
    {
        public int fieldId { get; set; }
        public int fieldType { get; set; }
        public string fieldName { get; set; }
        public List<UserCompany> customFieldValueOptions { get; set; }
    }

    public enum AtlassianGrouptypes { Jira, Confluence }

    public class AtlassianLicenceGroup
    {
        public AtlassianGrouptypes type { get; set; }
        public string groupName { get; set; }
        public AtlassianLicenceGroup(AtlassianGrouptypes _type, string _name)
        {
            type = _type;
            groupName = _name;
        }
    }

    public class Company
    {
        public int id { get; set; }
        public string name { get; set; }
        public int customFieldId { get; set; }
        public string addedByUserKey { get; set; }
        public string description { get; set; }
        public int UserCountJira { get; set; }
        public int UserCountConfluence { get; set; }
        public int UserCount { get; set; }
    }

    public class CompanyRootObject
    {
        public int id { get; set; }
        public int type { get; set; }
        public string name { get; set; }
        public bool displayInProjectTabPanel { get; set; }
        public bool displayInUserProfilePanel { get; set; }
        public bool usersMayCreateOptions { get; set; }
        public string description { get; set; }
        
        [JsonPropertyName("listOptions")]
        public List<Company> companies { get; set; }
    }

}
