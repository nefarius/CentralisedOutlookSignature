using System.Collections.Generic;
using System.DirectoryServices;
using log4net;

namespace CentralisedOutlookSignature.Scripts
{
    public class TagDefinitions
    {
        public IDictionary<string, string> Initialize(DirectoryEntry adUser, ILog log)
        {
            var adDisplayName = adUser.Properties["displayName"].Value as string;
            log.DebugFormat("adDisplayName = {0}", adDisplayName);

            var adEmailAddress = adUser.Properties["mail"].Value as string;
            log.DebugFormat("adEmailAddress = {0}", adEmailAddress);

            var adTitle = adUser.Properties["title"].Value;
            log.DebugFormat("adTitle = {0}", adTitle);

            var adDescription = adUser.Properties["description"].Value;
            log.DebugFormat("adDescription = {0}", adDescription);

            var adTelePhoneNumber = adUser.Properties["telephoneNumber"].Value as string;
            log.DebugFormat("adTelePhoneNumber = {0}", adTelePhoneNumber);

            var adMobile = adUser.Properties["mobile"].Value as string;
            log.DebugFormat("adMobile = {0}", adMobile);

            var adFaxNumber = adUser.Properties["facsimileTelephoneNumber"].Value as string;
            log.DebugFormat("adFaxNumber = {0}", adFaxNumber);

            var adDepartment = adUser.Properties["department"].Value as string;
            log.DebugFormat("adDepartment = {0}", adDepartment);

            var adHomePage = adUser.Properties["wWWHomePage"].Value as string;
            log.DebugFormat("adHomePage = {0}", adHomePage);

            var adCity = adUser.Properties["l"].Value as string;
            log.DebugFormat("adCity = {0}", adCity);

            var adStreetAddress = adUser.Properties["streetaddress"].Value as string;
            log.DebugFormat("adStreetAddress = {0}", adStreetAddress);

            var adPostalCode = adUser.Properties["postalCode"].Value as string;
            log.DebugFormat("adPostalCode = {0}", adPostalCode);
            
            return new Dictionary<string, string>
            {
                {"{Display name}", adDisplayName},
                {"{Department}", adDepartment},
                {"{Street}", adStreetAddress},
                {"{Postal code}", adPostalCode},
                {"{City}", adCity},
                {"{Phone}", adTelePhoneNumber},
                {"{Mobile}", adMobile},
                {"{Fax}", adFaxNumber},
                {"{e-mail}", adEmailAddress},
                {"{Web Page}", adHomePage}
            };
        }
    }
}