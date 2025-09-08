using System;
using System.Collections.Generic;
using System.DirectoryServices;

namespace ActiveDirectoryLookup.NetFramework
{
    /// <summary>
    /// Represents an Active Directory user with essential properties.
    /// This class is serializable to support cross-AppDomain and remoting scenarios.
    /// </summary>
    [Serializable]
    public class ActiveDirectoryUser
    {
        /// <summary>
        /// Gets or sets the employee ID of the user from Active Directory.
        /// This property maps to the 'employeeID' attribute in AD.
        /// </summary>
        public string EmployeeID { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the email address of the user.
        /// This property maps to the 'mail' attribute in Active Directory.
        /// </summary>
        public string Email { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the display name of the user.
        /// This property maps to the 'displayName' attribute in AD, falling back to 'cn' if displayName is not available.
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the Security Account Manager (SAM) account name of the user.
        /// This is the unique logon name used for Windows authentication.
        /// This property maps to the 'sAMAccountName' attribute in Active Directory.
        /// </summary>
        public string SamAccountName { get; set; } = string.Empty;
    }

    /// <summary>
    /// Represents the result of an Active Directory search operation.
    /// Contains the list of users found, success status, and error information.
    /// This class is serializable to support cross-AppDomain and remoting scenarios.
    /// </summary>
    [Serializable]
    public class SearchResult
    {
        /// <summary>
        /// Gets or sets the list of Active Directory users found during the search operation.
        /// </summary>
        public List<ActiveDirectoryUser> Users { get; set; } = new List<ActiveDirectoryUser>();

        /// <summary>
        /// Gets or sets a value indicating whether the search operation was successful.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Gets or sets the error message if the operation failed, or success details if the operation succeeded.
        /// </summary>
        public string ErrorMessage { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the total number of users found during the search operation.
        /// </summary>
        public int TotalFound { get; set; }
    }

    /// <summary>
    /// Provides Active Directory search and lookup functionality using System.DirectoryServices.
    /// This service is designed to run on .NET Framework 4.8 and act as an interface DLL 
    /// for client applications that need Active Directory access.
    /// 
    /// <para>
    /// <strong>Usage Instructions for Client Applications:</strong>
    /// </para>
    /// <para>
    /// 1. Add a reference to this DLL in your client project
    /// </para>
    /// <para>
    /// 2. Ensure the client application runs with appropriate Active Directory permissions
    /// </para>
    /// <para>
    /// 3. Create an instance of ActiveDirectoryService and call the desired methods
    /// </para>
    /// <para>
    /// 4. All methods return SearchResult objects with Success/ErrorMessage for error handling
    /// </para>
    /// 
    /// <example>
    /// <code>
    /// var adService = new ActiveDirectoryService();
    /// var result = adService.SearchUsersByName("John");
    /// if (result.Success)
    /// {
    ///     foreach (var user in result.Users)
    ///     {
    ///         Console.WriteLine($"{user.Name} - {user.Email}");
    ///     }
    /// }
    /// else
    /// {
    ///     Console.WriteLine($"Error: {result.ErrorMessage}");
    /// }
    /// </code>
    /// </example>
    /// </summary>
    public class ActiveDirectoryService
    {
        /// <summary>
        /// Searches for Active Directory users by name using wildcard matching.
        /// Searches across displayName, common name (cn), and SAM account name fields.
        /// </summary>
        /// <param name="searchName">
        /// The name or partial name to search for. Wildcard characters are automatically added.
        /// Case-insensitive search across multiple name fields.
        /// </param>
        /// <returns>
        /// A <see cref="SearchResult"/> containing:
        /// - List of matching users (limited to 100 results)
        /// - Success status indicating if the operation completed without errors
        /// - Error message if the operation failed, or success details if completed
        /// - Total count of users found
        /// </returns>
        /// <remarks>
        /// This method uses the current user's Windows credentials for Active Directory authentication.
        /// Requires appropriate read permissions on the Active Directory domain.
        /// Search is performed against the default domain context.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// var result = service.SearchUsersByName("Smith");
        /// 
        /// if (result.Success)
        /// {
        ///     Console.WriteLine($"Found {result.TotalFound} users");
        ///     foreach (var user in result.Users)
        ///     {
        ///         Console.WriteLine($"{user.Name} ({user.SamAccountName}) - {user.Email}");
        ///     }
        /// }
        /// </code>
        /// </example>
        public SearchResult SearchUsersByName(string searchName)
        {
            var result = new SearchResult();

            try
            {
                // Create a DirectoryEntry object for the current domain
                using (var domain = new DirectoryEntry())
                {
                    // Create a DirectorySearcher to search for users
                    using (var searcher = new DirectorySearcher(domain))
                    {
                        // Set the filter to search for users with the specified name
                        searcher.Filter = $"(&(objectClass=user)(objectCategory=person)(|(displayName=*{searchName}*)(cn=*{searchName}*)(sAMAccountName=*{searchName}*)))";

                        // Specify which properties to retrieve
                        searcher.PropertiesToLoad.Add("displayName");
                        searcher.PropertiesToLoad.Add("mail");
                        searcher.PropertiesToLoad.Add("employeeID");
                        searcher.PropertiesToLoad.Add("cn");
                        searcher.PropertiesToLoad.Add("sAMAccountName");

                        // Set search scope and size limit
                        searcher.SearchScope = SearchScope.Subtree;
                        searcher.SizeLimit = 100;

                        // Execute the search
                        using (var results = searcher.FindAll())
                        {
                            foreach (System.DirectoryServices.SearchResult searchResult in results)
                            {
                                var user = new ActiveDirectoryUser();

                                // Get display name or common name
                                if (searchResult.Properties["displayName"].Count > 0)
                                {
                                    user.Name = searchResult.Properties["displayName"][0]?.ToString() ?? string.Empty;
                                }
                                else if (searchResult.Properties["cn"].Count > 0)
                                {
                                    user.Name = searchResult.Properties["cn"][0]?.ToString() ?? string.Empty;
                                }

                                // Get email address
                                if (searchResult.Properties["mail"].Count > 0)
                                {
                                    user.Email = searchResult.Properties["mail"][0]?.ToString() ?? string.Empty;
                                }

                                // Get employee ID
                                if (searchResult.Properties["employeeID"].Count > 0)
                                {
                                    user.EmployeeID = searchResult.Properties["employeeID"][0]?.ToString() ?? string.Empty;
                                }

                                // Get SAM account name
                                if (searchResult.Properties["sAMAccountName"].Count > 0)
                                {
                                    user.SamAccountName = searchResult.Properties["sAMAccountName"][0]?.ToString() ?? string.Empty;
                                }

                                result.Users.Add(user);
                            }
                        }
                    }
                }

                result.Success = true;
                result.TotalFound = result.Users.Count;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error searching Active Directory: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Searches for a specific Active Directory user by their exact SAM account name.
        /// This method performs an exact match search for a single user.
        /// </summary>
        /// <param name="samAccountName">
        /// The exact SAM account name (Windows logon name) to search for.
        /// This should be the user's unique identifier in Active Directory.
        /// </param>
        /// <returns>
        /// A <see cref="SearchResult"/> containing:
        /// - The matching user (should be 0 or 1 user for exact matches)
        /// - Success status indicating if the operation completed without errors
        /// - Error message if the operation failed, or success details if completed
        /// - Total count of users found (typically 0 or 1)
        /// </returns>
        /// <remarks>
        /// This method is ideal for looking up a specific user when you know their exact logon name.
        /// Uses the current user's Windows credentials for Active Directory authentication.
        /// Search is limited to 1 result since SAM account names should be unique.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// var result = service.SearchBySamAccountName("jsmith");
        /// 
        /// if (result.Success && result.Users.Count > 0)
        /// {
        ///     var user = result.Users[0];
        ///     Console.WriteLine($"User: {user.Name}");
        ///     Console.WriteLine($"Email: {user.Email}");
        ///     Console.WriteLine($"Employee ID: {user.EmployeeID}");
        /// }
        /// else if (result.Success)
        /// {
        ///     Console.WriteLine("User not found");
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"Search failed: {result.ErrorMessage}");
        /// }
        /// </code>
        /// </example>
        public SearchResult SearchBySamAccountName(string samAccountName)
        {
            var result = new SearchResult();

            try
            {
                // Create a DirectoryEntry object for the current domain
                using (var domain = new DirectoryEntry())
                {
                    // Create a DirectorySearcher to search for users
                    using (var searcher = new DirectorySearcher(domain))
                    {
                        // Set the filter to search for specific SAM account name
                        searcher.Filter = $"(&(objectClass=user)(objectCategory=person)(sAMAccountName={samAccountName}))";

                        // Specify which properties to retrieve
                        searcher.PropertiesToLoad.Add("displayName");
                        searcher.PropertiesToLoad.Add("mail");
                        searcher.PropertiesToLoad.Add("employeeID");
                        searcher.PropertiesToLoad.Add("cn");
                        searcher.PropertiesToLoad.Add("sAMAccountName");

                        // Set search scope and size limit
                        searcher.SearchScope = SearchScope.Subtree;
                        searcher.SizeLimit = 1; // Should only find one user

                        // Execute the search
                        using (var results = searcher.FindAll())
                        {
                            foreach (System.DirectoryServices.SearchResult searchResult in results)
                            {
                                var user = new ActiveDirectoryUser();

                                // Get display name or common name
                                if (searchResult.Properties["displayName"].Count > 0)
                                {
                                    user.Name = searchResult.Properties["displayName"][0]?.ToString() ?? string.Empty;
                                }
                                else if (searchResult.Properties["cn"].Count > 0)
                                {
                                    user.Name = searchResult.Properties["cn"][0]?.ToString() ?? string.Empty;
                                }

                                // Get email address
                                if (searchResult.Properties["mail"].Count > 0)
                                {
                                    user.Email = searchResult.Properties["mail"][0]?.ToString() ?? string.Empty;
                                }

                                // Get employee ID
                                if (searchResult.Properties["employeeID"].Count > 0)
                                {
                                    user.EmployeeID = searchResult.Properties["employeeID"][0]?.ToString() ?? string.Empty;
                                }

                                // Get SAM account name
                                if (searchResult.Properties["sAMAccountName"].Count > 0)
                                {
                                    user.SamAccountName = searchResult.Properties["sAMAccountName"][0]?.ToString() ?? string.Empty;
                                }

                                result.Users.Add(user);
                            }
                        }
                    }
                }

                result.Success = true;
                result.TotalFound = result.Users.Count;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error searching Active Directory by SAM account: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Tests the basic connectivity to Active Directory without performing user searches.
        /// This method verifies that the application can establish a connection to the domain.
        /// </summary>
        /// <returns>
        /// A <see cref="SearchResult"/> containing:
        /// - Empty Users list (no users are searched in this test)
        /// - Success status indicating if the connection was established
        /// - Domain information if successful, or error details if failed
        /// - TotalFound will be 0 (no users searched)
        /// </returns>
        /// <remarks>
        /// This is a lightweight test that only verifies domain connectivity.
        /// Use <see cref="TestConnectionAndSearchCurrentUser"/> for a more comprehensive test
        /// that also verifies search permissions.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// var result = service.TestConnection();
        /// 
        /// if (result.Success)
        /// {
        ///     Console.WriteLine("AD Connection successful");
        ///     Console.WriteLine(result.ErrorMessage); // Contains domain info
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"AD Connection failed: {result.ErrorMessage}");
        /// }
        /// </code>
        /// </example>
        public SearchResult TestConnection()
        {
            var result = new SearchResult();

            try
            {
                using (var domain = new DirectoryEntry())
                {
                    var path = domain.Path;
                    var name = domain.Name;

                    result.Success = true;
                    result.ErrorMessage = $"Connection successful. Domain: {name}, Path: {path}";
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection failed: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Performs a comprehensive test of Active Directory connectivity and search functionality.
        /// This method not only tests the connection but also verifies that user search operations work
        /// by attempting to find the current user's profile in Active Directory.
        /// </summary>
        /// <returns>
        /// A <see cref="SearchResult"/> containing:
        /// - The current user's profile if found (Users list may contain 0 or 1 user)
        /// - Success status indicating if both connection and search permissions are working
        /// - Detailed information about the connection, domain, and current user profile
        /// - Warning message if connection works but user search permissions are limited
        /// </returns>
        /// <remarks>
        /// This is the most comprehensive connectivity test available. It verifies:
        /// - Basic Active Directory connectivity
        /// - User search functionality and permissions  
        /// - Ability to retrieve user attributes (name, email, employee ID)
        /// - Current user's presence in Active Directory
        /// 
        /// This method is recommended for initial setup and troubleshooting scenarios.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// var result = service.TestConnectionAndSearchCurrentUser();
        /// 
        /// if (result.Success)
        /// {
        ///     Console.WriteLine("Full AD functionality verified");
        ///     Console.WriteLine(result.ErrorMessage); // Contains detailed info
        ///     
        ///     if (result.Users.Count > 0)
        ///     {
        ///         var currentUser = result.Users[0];
        ///         Console.WriteLine($"Current user: {currentUser.Name}");
        ///     }
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"AD test failed: {result.ErrorMessage}");
        /// }
        /// </code>
        /// </example>
        public SearchResult TestConnectionAndSearchCurrentUser()
        {
            var result = new SearchResult();

            try
            {
                using (var domain = new DirectoryEntry())
                {
                    var path = domain.Path;
                    var name = domain.Name;

                    // Get current user's SAM account name
                    var currentUserSam = Environment.UserName;

                    // Test search functionality by finding the current user
                    var userSearchResult = SearchBySamAccountName(currentUserSam);

                    if (userSearchResult.Success && userSearchResult.Users.Count > 0)
                    {
                        var currentUser = userSearchResult.Users[0];
                        result.Success = true;
                        result.Users = userSearchResult.Users; // Include the found user in the result
                        result.TotalFound = userSearchResult.TotalFound;
                        result.ErrorMessage = $"Connection successful. Domain: {name}, Path: {path}\n\n" +
                                             $"Current user profile found:\n" +
                                             $"Name: {currentUser.Name}\n" +
                                             $"SAM Account: {currentUser.SamAccountName}\n" +
                                             $"Email: {currentUser.Email}\n" +
                                             $"Employee ID: {currentUser.EmployeeID}\n\n" +
                                             $"User search permissions verified successfully.";
                    }
                    else
                    {
                        result.Success = true;
                        result.ErrorMessage = $"Connection successful. Domain: {name}, Path: {path}\n\n" +
                                             $"Warning: Could not find current user profile for '{currentUserSam}'.\n" +
                                             $"This may indicate limited search permissions or the user is not in Active Directory.\n" +
                                             $"Error: {userSearchResult.ErrorMessage}";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection test failed: {ex.Message}";
            }

            return result;
        }
    }
}