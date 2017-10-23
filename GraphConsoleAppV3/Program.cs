#region

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

#endregion

namespace GraphConsoleAppV3
{
    public class Program
    {
        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();

            #region Setup Active Directory Client

            //*********************************************************************
            // setup Active Directory Client
            //*********************************************************************
            ActiveDirectoryClient activeDirectoryClient;
            try
            {
                activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            }
            catch (AuthenticationException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            #endregion

            #region TenantDetails

            //*********************************************************************
            // Get Tenant Details
            // Note: update the string TenantId with your TenantId.
            // This can be retrieved from the login Federation Metadata end point:             
            // https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml
            //  Replace "GraphDir1.onMicrosoft.com" with any domain owned by your organization
            // The returned value from the first xml node "EntityDescriptor", will have a STS URL
            // containing your TenantId e.g. "https://sts.windows.net/4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34/" is returned for GraphDir1.onMicrosoft.com
            //*********************************************************************
            VerifiedDomain initialDomain = new VerifiedDomain();
            VerifiedDomain defaultDomain = new VerifiedDomain();
            ITenantDetail tenant = null;
            Console.WriteLine("\n Retrieving Tenant Details");
            try
            {
                List<ITenantDetail> tenantsList = activeDirectoryClient.TenantDetails
                    .Where(tenantDetail => tenantDetail.ObjectId.Equals(Constants.TenantId))
                    .ExecuteAsync().Result.CurrentPage.ToList();
                if (tenantsList.Count > 0)
                {
                    tenant = tenantsList.First();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting TenantDetails {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (tenant == null)
            {
                Console.WriteLine("Tenant not found");
            }
            else
            {
                TenantDetail tenantDetail = (TenantDetail)tenant;
                Console.WriteLine("Tenant Display Name: " + tenantDetail.DisplayName);

                // Get the Tenant's Verified Domains 
                initialDomain = tenantDetail.VerifiedDomains.First(x => x.Initial.HasValue && x.Initial.Value);
                Console.WriteLine("Initial Domain Name: " + initialDomain.Name);
                defaultDomain = tenantDetail.VerifiedDomains.First(x => x.@default.HasValue && x.@default.Value);
                Console.WriteLine("Default Domain Name: " + defaultDomain.Name);

                // Get Tenant's Tech Contacts
                foreach (string techContact in tenantDetail.TechnicalNotificationMails)
                {
                    Console.WriteLine("Tenant Tech Contact: " + techContact);
                }
            }

            #endregion

            #region Create a new User




            IUser newUser = new User();
            if (defaultDomain.Name != null)
            {
                newUser.DisplayName = "demo1";
                newUser.UserPrincipalName = "xuhua00101" + "@" + defaultDomain.Name;
                newUser.AccountEnabled = true;
                newUser.MailNickname = "SampleAppDemoUserManager";
                newUser.PasswordPolicies = "DisablePasswordExpiration";
                newUser.PasswordProfile = new PasswordProfile
                {
                    Password = "TempP@ssw0rd!",
                    ForceChangePasswordNextLogin = true
                };
                newUser.UsageLocation = "US";
                try
                {
                    //activeDirectoryClient.Users.AddUserAsync(newUser).Wait();
                    Console.WriteLine("\nNew User {0} was created", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError creating new user {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }


            #endregion

           // IApplication APPl = new Application();
           //// APPl.AppId = "12345678";
           // // APPl.AppRoles = 
           // APPl.AvailableToOtherTenants = false;
           // APPl.DisplayName = "vik00de";
           //// APPl.IdentifierUris
           // APPl.Homepage = "https://ww.baidu.com";
           // // APPl.IdentifierUris = "https://ww.baidu.com1";
           // APPl.LogoutUrl = "https://ww.baidu.com";
           // APPl.ErrorUrl = "https://ww.baidu.com1/1";
           //// IList<string> ls = APPl.IdentifierUris;
           //APPl.IdentifierUris.Add("https://localhost/demo/" + Guid.NewGuid());

            Application newApp = new Application { DisplayName = "wode" + Helper.GetRandomString(8) };
            newApp.IdentifierUris.Add("https://localhost/demo1/" + Guid.NewGuid());
            newApp.ReplyUrls.Add("https://localhost/demo1");
            newApp.PublicClient = null;
            AppRole appRole = new AppRole()
            { 
                Id = Guid.NewGuid(),
                IsEnabled = true,
                DisplayName = "Something",
                Description = "Anything",
                Value = "policy.write"


            };
            appRole.AllowedMemberTypes.Add("User");
            newApp.AppRoles.Add(appRole);

           


            //AzureADServicePrincipal
            //PasswordCredential password = new PasswordCredential
            //{
            //    StartDate = DateTime.UtcNow,
            //    EndDate = DateTime.UtcNow.AddYears(1),
            //    Value = "password",
            //    KeyId = Guid.NewGuid()
            //};
            //newApp.PasswordCredentials.Add(password);
            try
            {
                
                 activeDirectoryClient.Applications.AddApplicationAsync(newApp).Wait();
                Console.WriteLine("New Application created: " + newApp.DisplayName);
            }
            catch (Exception e)
            {
                string a = e.Message.ToString();
               // Program.WriteError("\nError ceating Application: {0}", Program.ExtractErrorMessage(e));
            }

            ServicePrincipal s = new ServicePrincipal();
            s.Tags.Add("WindowsAzureActiveDirectoryIntegratedApp");
            s.AppId = newApp.AppId;
            try
            {
                activeDirectoryClient.ServicePrincipals.AddServicePrincipalAsync(s).Wait();
            }
            catch (Exception e) {
                string a = e.Message.ToString();
            }

            //try
            //{
            //    activeDirectoryClient.Applications.AddApplicationAsync(appObject).Wait();
            //}
            //catch (Exception e) {
            //    string mess = e.Message.ToString();
            //    string a = "";
            //}

            #region Create a User with a temp Password

            //*********************************************************************************************
            // Create a new User with a temp Password
            //*********************************************************************************************
            IUser userToBeAdded = new User();
            
            userToBeAdded.DisplayName = "Sample App Demo User";
            userToBeAdded.UserPrincipalName = Helper.GetRandomString(10) + "@" + defaultDomain.Name;
            userToBeAdded.AccountEnabled = true;
            userToBeAdded.MailNickname = "SampleAppDemoUser";


            userToBeAdded.PasswordProfile = new PasswordProfile
            {
                Password = "TempP@ssw0rd!",
                ForceChangePasswordNextLogin = true
            };
            userToBeAdded.UsageLocation = "US";
            try
            {
                activeDirectoryClient.Users.AddUserAsync(userToBeAdded).Wait();
                Console.WriteLine("\nNew User {0} was created", userToBeAdded.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError creating new user. {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
          //  activeDirectoryClient.Applications.AddApplicationAsync(iApp);
            #endregion

            #region Create a new Group

            //*********************************************************************************************
            // Create a new Group
            //*********************************************************************************************
            Group californiaEmployees = new Group
            {
                DisplayName = "California Employees" + Helper.GetRandomString(8),
                Description = "Employees in the state of California",
                MailNickname = "CalEmployees",
                MailEnabled = false,
                SecurityEnabled = true
                
            };
            try
            {
                activeDirectoryClient.Groups.AddGroupAsync(californiaEmployees).Wait();
                Console.WriteLine("\nNew Group {0} was created", californiaEmployees.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError creating new Group {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Search for Group using StartWith filter

            //*********************************************************************
            // Search for a group using a startsWith filter (displayName property)
            //*********************************************************************
            Group retrievedGroup = new Group();
            string searchString = "California Employees";
            List<IGroup> foundGroups = null;
            try
            {
                foundGroups = activeDirectoryClient.Groups
                    .Where(group => group.DisplayName.StartsWith(searchString))
                    .ExecuteAsync().Result.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Group {0} {1}",
                    e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (foundGroups != null && foundGroups.Count > 0)
            {
                retrievedGroup = foundGroups.First() as Group;
            }
            else
            {
                Console.WriteLine("Group Not Found");
            }

            #endregion

            #region Assign Member to Group

            if (retrievedGroup.ObjectId != null)
            {
                try
                {
                    retrievedGroup.Members.Add(newUser as DirectoryObject);
                    retrievedGroup.UpdateAsync().Wait();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError assigning member to group. {0} {1}",
                        e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }

            }

            #endregion


            #region Add User to Group

            //*********************************************************************************************
            // Add User to the "WA" Group 
            //*********************************************************************************************
            if (retrievedGroup.ObjectId != null)
            {
                try
                {
                    retrievedGroup.Members.Add(userToBeAdded as DirectoryObject);
                    retrievedGroup.UpdateAsync().Wait();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nAdding user to group failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion


            #region Get Group members

            if (retrievedGroup.ObjectId != null)
            {
                Console.WriteLine("\n Found Group: " + retrievedGroup.DisplayName + "  " + retrievedGroup.Description);

                //*********************************************************************
                // get the groups' membership - 
                // Note this method retrieves ALL links in one request - please use this method with care - this
                // may return a very large number of objects
                //*********************************************************************
                IGroupFetcher retrievedGroupFetcher = retrievedGroup;
                try
                {
                    IPagedCollection<IDirectoryObject> members = retrievedGroupFetcher.Members.ExecuteAsync().Result;
                    Console.WriteLine(" Members:");
                    do
                    {
                        List<IDirectoryObject> directoryObjects = members.CurrentPage.ToList();
                        foreach (IDirectoryObject member in directoryObjects)
                        {
                            if (member is User)
                            {
                                User user = member as User;
                                Console.WriteLine("User DisplayName: {0}  UPN: {1}",
                                    user.DisplayName,
                                    user.UserPrincipalName);
                            }
                            if (member is Group)
                            {
                                Group group = member as Group;
                                Console.WriteLine("Group DisplayName: {0}", group.DisplayName);
                            }
                            if (member is Contact)
                            {
                                Contact contact = member as Contact;
                                Console.WriteLine("Contact DisplayName: {0}", contact.DisplayName);
                            }
                        }
                        members = members.GetNextPageAsync().Result;
                    } while (members != null);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError getting groups' membership. {0} {1}",
                        e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion



            
            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", currentDateTime);
            // Console.ReadKey();
            Console.WriteLine();
            #region Search User by UPN

            // search for a single user by UPN
            searchString = "" + "@" + initialDomain.Name;
            Console.WriteLine("\n Retrieving user with UPN {0}", searchString);
            User retrievedUser = new User();
            List<IUser> retrievedUsers = null;
            try
            {
                retrievedUsers = activeDirectoryClient.Users
                    .Where(user => user.UserPrincipalName.Equals(searchString))
                    .ExecuteAsync().Result.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting new user {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            // should only find one user with the specified UPN
            if (retrievedUsers != null && retrievedUsers.Count == 1)
            {
                retrievedUser = (User)retrievedUsers.First();
            }
            else
            {
                Console.WriteLine("User not found {0}", searchString);
            }

            #endregion
            Console.WriteLine("\n {0} is a member of the following Group and Roles (IDs)", retrievedUser.DisplayName);
            IUserFetcher retrievedUserFetcher = retrievedUser;
            try
            {
                IPagedCollection<IDirectoryObject> pagedCollection = retrievedUserFetcher.MemberOf.ExecuteAsync().Result;
                do
                {
                    List<IDirectoryObject> directoryObjects = pagedCollection.CurrentPage.ToList();
                    foreach (IDirectoryObject directoryObject in directoryObjects)
                    {
                        if (directoryObject is Group)
                        {
                            Group group = directoryObject as Group;
                            Console.WriteLine(" Group: {0}  Description: {1}", group.DisplayName, group.Description);
                        }
                        if (directoryObject is DirectoryRole)
                        {
                            DirectoryRole role = directoryObject as DirectoryRole;
                            Console.WriteLine(" Role: {0}  Description: {1}", role.DisplayName, role.Description);
                        }
                    }
                    pagedCollection = pagedCollection.GetNextPageAsync().Result;
                } while (pagedCollection != null);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting user's groups and roles memberships. {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            Console.WriteLine("Press any key to continue....");
            Console.ReadKey();
        }
    }
}