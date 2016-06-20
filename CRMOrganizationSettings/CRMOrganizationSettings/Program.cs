using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace CRMOrganizationSettings
{
    class Program
    {
        static void Main(string[] args)
        {
            CrmServiceClient crmSvc = new CrmServiceClient("yourUserId", CrmServiceClient.MakeSecureString("yourPassword"), "NorthAmerica", "UniqueOrgName", isOffice365: true);

            using (OrganizationServiceProxy proxy = crmSvc.OrganizationServiceProxy)
            {
                EntityCollection organizationSettings;
                bool IsSOPIntegrationEnabled;
                QueryForSOPSetting(proxy, out organizationSettings, out IsSOPIntegrationEnabled);

                Console.WriteLine(string.Format("IsSOPIntegrationEnabled: {0}", IsSOPIntegrationEnabled.ToString()));

                if (IsSOPIntegrationEnabled)
                    IsSOPIntegrationEnabled = false;

                Entity org = new Entity("organization");
                org.Id = organizationSettings.Entities[0].Id;
                org["issopintegrationenabled"] = false;

                proxy.Update(org);

                Console.WriteLine(string.Format("IsSOPIntegrationEnabled: {0}", IsSOPIntegrationEnabled.ToString()));

            }
        }

        private static void QueryForSOPSetting(OrganizationServiceProxy proxy, out EntityCollection organizationSettings, out bool IsSOPIntegrationEnabled)
        {
            QueryExpression query = new QueryExpression()
            {
                EntityName = "organization",
                ColumnSet = new ColumnSet(true),
            };

            organizationSettings = proxy.RetrieveMultiple(query);

            IsSOPIntegrationEnabled = organizationSettings.Entities[0].GetAttributeValue<Boolean>("issopintegrationenabled");
        }
    }
}
