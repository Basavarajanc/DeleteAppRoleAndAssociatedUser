using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using OfficeOpenXml;
using AuthenticationException = System.Security.Authentication.AuthenticationException;

namespace ListAndDeleteAppRoles
{
    public class Program
    {
        private static GraphServiceClient? _graphClient;
        private static string? _accessToken;
        private static readonly string? ClientId = "<>";
        private static readonly string? Tenant = "<>";
        private static readonly string? ClientSecret = "<>";
        private static readonly string? ObjectId = "<>";
        private static readonly string? AzureServicePrinciple = "<>";
        private static string? SearchString = "TestBasav";

        private static async Task Main()
        {
            /*//List the users with Empty Assigned roles
          var result =  await ListUsersAssignedToRole();
          var countAppRoles = 0;
          
          var appUsersData = new List<AppUserData>();
          AppUserData appUserData;
          
          foreach (var appRole in result)
          {
              countAppRoles++;
              appUserData = new AppUserData();
              appUserData.roleNo = countAppRoles.ToString();
             //appUserData.resourceDisplayName = appRole.ResourceDisplayName;
              //appUserData.principleDisplayName = appRole.PrincipalDisplayName;
              //appUserData.appRoleId = appRole.AppRoleId.ToString();
              //appUserData.resourceId = appRole.ResourceId.ToString();
             // appUserData.deleteDateTime = appRole.DeletedDateTime.ToString();
              
              appUsersData.Add(appUserData);

          }

          ExportToExcel(appUsersData);*/

          //Delete the Users
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
            foreach (var appRole in appRoles)
            {
                if ((bool)appRole?.Description.Contains(SearchString))
                {
                    var isDisabledAppRole = await DisableAppRole(appRole.Id.Value);
      
                    if (isDisabledAppRole)
                    {
                        var removedAppRole = await RemoveAppRoleFromUsers(appRole.Id.Value);
                
                        if (removedAppRole)
                        {
                            await DeleteAppRole(appRole.Id.Value);
                        }
                    }
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
            try
            {
                var usersAssignedToRole = await GetUsersAssignedToRole(id);
                foreach (var userAssignment in usersAssignedToRole)
                    await _graphClient?.ServicePrincipals[AzureServicePrinciple].AppRoleAssignedTo[userAssignment.Id]
                        .Request()
                        .DeleteAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
           
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
        
        public static async Task<List<User>> ListUsersAssignedToRole()
        {
            var app = await GetAzureApp();
            var usersWithoutAppRoleAssignments = new List<User>();
            try
            {
                var appRoleAssignments = new List<AppRoleAssignment>();
                var usersList = new List<User>();

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

                // Get users assigned to the app
                List<User> assignedUsers = new List<User>();
                foreach (var assignment in appRoleAssignments)
                {
                    if (assignment.PrincipalType == "User")
                    {
                        var user = await _graphClient.Users[assignment.PrincipalId.ToString()]
                            .Request()
                            .GetAsync();

                        assignedUsers.Add(user);
                    }
                }
               
            }
            catch (Exception ex)
            {
                throw;
            }
            return usersWithoutAppRoleAssignments;
        }
        
        private static async Task ExportToExcel(List<AppUserData> data)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    // Add a worksheet
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Add headers
                    worksheet.Cells["A1"].Value = "RoleNo";
                    worksheet.Cells["B1"].Value = "resourceDisplayName";
                    worksheet.Cells["C1"].Value = "principleDisplayName";
                    worksheet.Cells["D1"].Value = "appRoleId";
                    worksheet.Cells["E1"].Value = "deleteDateTime";
                    //worksheet.Cells["F1"].Value = "resourceId";

                    // Add data
                    for (var i = 0; i < data.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = data[i].roleNo;
                        worksheet.Cells[i + 2, 2].Value = data[i].resourceDisplayName;
                        worksheet.Cells[i + 2, 3].Value = data[i].principleDisplayName;
                        worksheet.Cells[i + 2, 4].Value = data[i].appRoleId;
                        worksheet.Cells[i + 2, 5].Value = data[i].deleteDateTime;
                       // worksheet.Cells[i + 2, 6].Value = data[i].resourceId;
                    }

                    // Save the file
                    var fileInfo = new FileInfo("output.xlsx");
                    package.SaveAs(fileInfo);
                }
            }
            catch (Exception ex)
            {
            }
        }
    }

    public class AppUserData
    {
        public string roleNo { get; set; }

        public string resourceDisplayName { get; set; }

        public string principleDisplayName { get; set; }

        public string appRoleId { get; set; }

        public string deleteDateTime { get; set; }

        public string resourceId { get; set; }
        
        public string userRoleAssignedTo { get; set; }
    }
}