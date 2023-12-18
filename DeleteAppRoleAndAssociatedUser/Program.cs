using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using AuthenticationException = System.Security.Authentication.AuthenticationException;

namespace ListAndDeleteAppRoles
{
    public class Program
    {
        private static GraphServiceClient? _graphClient;
        private static string? _accessToken;
        private static readonly string? ClientId = "bed8937e-3fa9-4d37-8bfb-aa0fca76c907";
        private static readonly string? Tenant = "fa74eeb7-373a-4c5b-8c97-4d330cfa9f60";
        private static readonly string? ClientSecret = "3br8Q~UBMbzFpd7xy0ZnN49zkZkiI7_s7YPAZcXK";
        private static readonly string? ObjectId = "4adb256d-90a6-492c-bcfb-746a1dacbf18";
        private static readonly string? AzureServicePrinciple = "8f2b53ec-192f-48bb-8b8e-330e4ecd4ee0";

        private static async Task Main()
        {
            await HandleRemoveAzureAppRole();
        }

        private static async Task<string> GetAccessToken()
        {
            var scopes = new List<string> { "https://graph.microsoft.com/.default" };

            var msalClient = ConfidentialClientApplicationBuilder
                .Create(ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, Tenant)
                .WithClientSecret(ClientSecret)
                .Build();

            var userAssertion = new UserAssertion("testDrive", "client_credentials");
            try
            {
                var token = await msalClient.AcquireTokenOnBehalfOf(scopes, userAssertion).ExecuteAsync();
                return token.AccessToken;
            }
            catch
            {
                throw new AuthenticationException("Issue on getting access token");
            }
        }

        private static GraphServiceClient? InitiateGraphServiceClient(string accessToken)
        {
            var authenticationProvider = new DelegateAuthenticationProvider(
                requestMessage =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    return Task.FromResult(0);
                });
            return new GraphServiceClient(authenticationProvider);
        }

        private static async Task<Application> GetAzureApp()
        {
            _accessToken = await GetAccessToken();
            _graphClient = InitiateGraphServiceClient(_accessToken);
            var application = await _graphClient.Applications[ObjectId].Request().GetAsync();

            return application;
        }

        private static async Task<bool> HandleRemoveAzureAppRole()
        {
            var app = await GetAzureApp();
            var appRoles = app.AppRoles.ToList();

            Guid guid = Guid.NewGuid();
            
            var isDisabledAppRole = await DisableAppRole(guid);

            if (isDisabledAppRole)
            {
                var removedAppRole = await RemoveAppRoleFromUsers(guid);
                
                if (removedAppRole)
                {
                    await DeleteAppRole(guid);
                }
            }

            return false;
        }

        private static async Task<bool> DisableAppRole(Guid hubId)
        {
            var app = await GetAzureApp();
            var appRoles = app.AppRoles.ToList();
            var modifiedAppRole = appRoles.FirstOrDefault(p => p.Id.Value.Equals(hubId));
            if (modifiedAppRole != null)
            {
                appRoles.Where(p => p.Id == modifiedAppRole.Id).ToList().ForEach(p =>
                {
                    p.Value = $"#HUB_{hubId}";
                    p.IsEnabled = false;
                    p.AllowedMemberTypes = new[] { "User" };
                });

                try
                {
                    app.AppRoles = appRoles;
                    //first save that the appRole is disabled 
                    await _graphClient.Applications[ObjectId]
                        .Request()
                        .UpdateAsync(app);

                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                    throw;
                }
            }

            return true;
        }

        private static async Task<bool> DeleteAppRole(Guid hubId)
        {
            var app = await GetAzureApp();
            var appRoles = app.AppRoles.ToList();
            var modifiedAppRole = appRoles.FirstOrDefault(p => p.Id.Value.Equals(hubId));

            // remove appRole
            if (modifiedAppRole != null)
            {
                appRoles.Remove(modifiedAppRole);
                app.AppRoles = appRoles;

                await _graphClient.Applications[ObjectId]
                    .Request()
                    .UpdateAsync(app);
            }

            return true;
        }

        private static async Task<bool> RemoveAppRoleFromUsers(Guid id)
        {
            var usersAssignedToRole = await GetUsersAssignedToRole(id);
            foreach (var userAssignment in usersAssignedToRole)
                await _graphClient?.ServicePrincipals[AzureServicePrinciple].AppRoleAssignedTo[userAssignment.Id]
                    .Request()
                    .DeleteAsync();
            return true;
        }

        public static async Task<List<AppRoleAssignment>> GetUsersAssignedToRole(Guid id)
        {
            var appRoleAssignments = new List<AppRoleAssignment>();

            var assignments = await _graphClient.ServicePrincipals[AzureServicePrinciple].AppRoleAssignedTo
                .Request()
                .Top(998)
                .GetAsync();

            appRoleAssignments.AddRange(assignments.CurrentPage);
            while (assignments.NextPageRequest != null)
            {
                assignments = await assignments.NextPageRequest.GetAsync();
                appRoleAssignments.AddRange(assignments.CurrentPage);
            }

            var result = appRoleAssignments.Where(x => x.AppRoleId == id).ToList();
            return result;
        }
    }

    public class AppRolesData
    {
        public string roleNo { get; set; }

        public string displayName { get; set; }

        public string description { get; set; }

        public string isEnabled { get; set; }

        public string value { get; set; }

        public string usersAssignedToRole { get; set; }
    }
}