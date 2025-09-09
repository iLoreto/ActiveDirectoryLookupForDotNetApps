using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;

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

        /// <summary>
        /// Gets or sets a value indicating whether the user is a supervisor.
        /// This is determined by checking membership in supervisor-related Active Directory groups.
        /// </summary>
        public bool IsSupervisor { get; set; } = false;

        /// <summary>
        /// Gets or sets a value indicating whether the user is a manager.
        /// This is determined by checking membership in manager-related Active Directory groups.
        /// </summary>
        public bool IsManager { get; set; } = false;
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
    /// Represents the result of a supervisor or manager membership check operation.
    /// </summary>
    [Serializable]
    public class RoleCheckResult
    {
        /// <summary>
        /// Gets or sets a value indicating whether the operation was successful.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user is a supervisor.
        /// </summary>
        public bool IsSupervisor { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user is a manager.
        /// </summary>
        public bool IsManager { get; set; }

        /// <summary>
        /// Gets or sets the error message if the operation failed, or details about role groups if succeeded.
        /// </summary>
        public string ErrorMessage { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the list of supervisor groups the user belongs to.
        /// </summary>
        public List<string> SupervisorGroups { get; set; } = new List<string>();

        /// <summary>
        /// Gets or sets the list of manager groups the user belongs to.
        /// </summary>
        public List<string> ManagerGroups { get; set; } = new List<string>();
    }

    /// <summary>
    /// Represents the result of a supervisor membership check operation.
    /// This class is kept for backward compatibility.
    /// </summary>
    [Serializable]
    public class SupervisorCheckResult
    {
        /// <summary>
        /// Gets or sets a value indicating whether the operation was successful.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user is a supervisor.
        /// </summary>
        public bool IsSupervisor { get; set; }

        /// <summary>
        /// Gets or sets the error message if the operation failed, or details about supervisor groups if succeeded.
        /// </summary>
        public string ErrorMessage { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the list of supervisor groups the user belongs to.
        /// </summary>
        public List<string> SupervisorGroups { get; set; } = new List<string>();
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
        /// Supports both "First Last" and "Last, First" name formats.
        /// </summary>
        /// <param name="searchName">
        /// The name or partial name to search for. Wildcard characters are automatically added.
        /// Case-insensitive search across multiple name fields.
        /// Supports multiple name formats (e.g., "John Smith" will find both "John Smith" and "Smith, John").
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
        /// Enhanced to handle multiple name formats for better search results.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// var result = service.SearchUsersByName("John Smith"); // Will find "John Smith" or "Smith, John"
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
                        // Build enhanced filter that handles multiple name formats
                        var filter = BuildNameSearchFilter(searchName);
                        searcher.Filter = filter;

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
        /// Builds an enhanced LDAP filter that handles multiple name formats.
        /// Supports both "First Last" and "Last, First" search patterns.
        /// </summary>
        /// <param name="searchName">The name to search for.</param>
        /// <returns>An LDAP filter string that searches multiple name formats.</returns>
        private string BuildNameSearchFilter(string searchName)
        {
            if (string.IsNullOrWhiteSpace(searchName))
            {
                return "(&(objectClass=user)(objectCategory=person))";
            }

            var searchTerms = new List<string>();
            
            // Add the original search term
            searchTerms.Add($"*{searchName.Trim()}*");

            // Check if the search contains multiple words (likely first and last name)
            var words = searchName.Trim().Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (words.Length == 2)
            {
                var firstWord = words[0];
                var secondWord = words[1];
                
                // Add "Last, First" format
                searchTerms.Add($"*{secondWord}, {firstWord}*");
                
                // Add "First*Last" format (handles middle names/initials)
                searchTerms.Add($"*{firstWord}*{secondWord}*");
                
                // Add "Last*First" format (handles middle names/initials in reverse)
                searchTerms.Add($"*{secondWord}*{firstWord}*");
            }
            else if (words.Length > 2)
            {
                // For more than 2 words, try first and last word combinations
                var firstWord = words[0];
                var lastWord = words[words.Length - 1];
                
                // Add "Last, First" format
                searchTerms.Add($"*{lastWord}, {firstWord}*");
                
                // Add partial matches for each word
                foreach (var word in words)
                {
                    if (word.Length > 2) // Only add words longer than 2 characters
                    {
                        searchTerms.Add($"*{word}*");
                    }
                }
            }

            // Build the OR conditions for displayName
            var displayNameConditions = string.Join("", searchTerms.Select(term => $"(displayName={term})"));
            
            // Build the OR conditions for cn (common name)
            var cnConditions = string.Join("", searchTerms.Select(term => $"(cn={term})"));
            
            // Build the OR conditions for sAMAccountName
            var samConditions = string.Join("", searchTerms.Select(term => $"(sAMAccountName={term})"));

            // Combine all conditions
            var filter = $"(&(objectClass=user)(objectCategory=person)(|{displayNameConditions}{cnConditions}{samConditions}))";
            
            return filter;
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
        /// Checks if a user is a supervisor or manager by examining their Active Directory group memberships.
        /// Searches for groups that contain supervisor, manager, or leadership-related keywords in their names.
        /// </summary>
        /// <param name="samAccountName">
        /// The SAM account name of the user to check for supervisor/manager privileges.
        /// </param>
        /// <returns>
        /// A <see cref="RoleCheckResult"/> containing:
        /// - Success status indicating if the operation completed without errors
        /// - IsSupervisor boolean indicating if the user has supervisor privileges
        /// - IsManager boolean indicating if the user has manager privileges
        /// - List of supervisor and manager groups the user belongs to
        /// - Error message if the operation failed, or group details if succeeded
        /// </returns>
        /// <remarks>
        /// This method searches for group memberships that typically indicate supervisory or managerial roles.
        /// The search is case-insensitive and looks for common role-related keywords in group names.
        /// Uses the current user's Windows credentials for Active Directory authentication.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// var result = service.CheckUserRoles("jsmith");
        /// 
        /// if (result.Success)
        /// {
        ///     if (result.IsSupervisor)
        ///     {
        ///         Console.WriteLine("User is a supervisor");
        ///         Console.WriteLine($"Supervisor groups: {string.Join(", ", result.SupervisorGroups)}");
        ///     }
        ///     if (result.IsManager)
        ///     {
        ///         Console.WriteLine("User is a manager");
        ///         Console.WriteLine($"Manager groups: {string.Join(", ", result.ManagerGroups)}");
        ///     }
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"Check failed: {result.ErrorMessage}");
        /// }
        /// </code>
        /// </example>
        public RoleCheckResult CheckUserRoles(string samAccountName)
        {
            var result = new RoleCheckResult();

            try
            {
                // Create a DirectoryEntry object for the current domain
                using (var domain = new DirectoryEntry())
                {
                    // Create a DirectorySearcher to find the user
                    using (var searcher = new DirectorySearcher(domain))
                    {
                        // Set the filter to search for the specific user
                        searcher.Filter = $"(&(objectClass=user)(objectCategory=person)(sAMAccountName={samAccountName}))";
                        
                        // Specify which properties to retrieve
                        searcher.PropertiesToLoad.Add("memberOf");
                        searcher.PropertiesToLoad.Add("displayName");
                        
                        // Set search scope and size limit
                        searcher.SearchScope = SearchScope.Subtree;
                        searcher.SizeLimit = 1;

                        // Execute the search
                        using (var results = searcher.FindAll())
                        {
                            if (results.Count == 0)
                            {
                                result.Success = true;
                                result.IsSupervisor = false;
                                result.IsManager = false;
                                result.ErrorMessage = $"User '{samAccountName}' not found in Active Directory.";
                                return result;
                            }

                            var userResult = results[0];
                            var displayName = userResult.Properties["displayName"].Count > 0 
                                ? userResult.Properties["displayName"][0]?.ToString() ?? samAccountName
                                : samAccountName;

                            // Define keywords that typically indicate supervisor roles
                            var supervisorKeywords = new[] 
                            { 
                                "supervisor", "lead", "team lead", "senior", "coordinator"
                            };

                            // Define keywords that typically indicate manager roles
                            var managerKeywords = new[] 
                            { 
                                "manager", "admin", "director", "chief", "head"
                            };

                            var supervisorGroups = new List<string>();
                            var managerGroups = new List<string>();

                            // Check group memberships
                            if (userResult.Properties["memberOf"].Count > 0)
                            {
                                foreach (string groupDN in userResult.Properties["memberOf"])
                                {
                                    // Extract group name from DN (Distinguished Name)
                                    var groupName = ExtractGroupNameFromDN(groupDN);
                                    
                                    // Check if group name contains supervisor keywords
                                    foreach (var keyword in supervisorKeywords)
                                    {
                                        if (groupName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            supervisorGroups.Add(groupName);
                                            break; // Avoid adding the same group multiple times
                                        }
                                    }

                                    // Check if group name contains manager keywords
                                    foreach (var keyword in managerKeywords)
                                    {
                                        if (groupName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            managerGroups.Add(groupName);
                                            break; // Avoid adding the same group multiple vezes
                                        }
                                    }
                                }
                            }

                            result.Success = true;
                            result.IsSupervisor = supervisorGroups.Count > 0;
                            result.IsManager = managerGroups.Count > 0;
                            result.SupervisorGroups = supervisorGroups;
                            result.ManagerGroups = managerGroups;

                            var roleInfo = new List<string>();
                            if (result.IsSupervisor)
                                roleInfo.Add($"Supervisor groups: {string.Join(", ", supervisorGroups)}");
                            if (result.IsManager)
                                roleInfo.Add($"Manager groups: {string.Join(", ", managerGroups)}");

                            if (result.IsSupervisor || result.IsManager)
                            {
                                result.ErrorMessage = $"User '{displayName}' has leadership roles.\n" + string.Join("\n", roleInfo);
                            }
                            else
                            {
                                result.ErrorMessage = $"User '{displayName}' has no supervisor or manager roles.\n" +
                                                     "No leadership-related group memberships found.";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.IsSupervisor = false;
                result.IsManager = false;
                result.ErrorMessage = $"Error checking user roles: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Extracts the group name from a Distinguished Name (DN) string.
        /// </summary>
        /// <param name="groupDN">The Distinguished Name of the group.</param>
        /// <returns>The extracted group name.</returns>
        private string ExtractGroupNameFromDN(string groupDN)
        {
            if (string.IsNullOrEmpty(groupDN))
                return string.Empty;

            // Group DN format: CN=GroupName,OU=...,DC=...,DC=...
            // Extract the CN part
            var cnStart = groupDN.IndexOf("CN=", StringComparison.OrdinalIgnoreCase);
            if (cnStart == -1)
                return groupDN;

            cnStart += 3; // Skip "CN="
            var cnEnd = groupDN.IndexOf(',', cnStart);
            
            if (cnEnd == -1)
                return groupDN.Substring(cnStart);
            
            return groupDN.Substring(cnStart, cnEnd - cnStart);
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
                        
                        // Check roles for the current user
                        var roleResult = CheckUserRoles(currentUserSam);
                        if (roleResult.Success)
                        {
                            currentUser.IsSupervisor = roleResult.IsSupervisor;
                            currentUser.IsManager = roleResult.IsManager;
                        }
                        
                        result.Success = true;
                        result.Users = userSearchResult.Users; // Include the found user in the result
                        result.TotalFound = userSearchResult.TotalFound;
                        result.ErrorMessage = $"Connection successful. Domain: {name}, Path: {path}\n\n" +
                                             $"Current user profile found:\n" +
                                             $"Name: {currentUser.Name}\n" +
                                             $"SAM Account: {currentUser.SamAccountName}\n" +
                                             $"Email: {currentUser.Email}\n" +
                                             $"Employee ID: {currentUser.EmployeeID}\n" +
                                             $"Is Supervisor: {(currentUser.IsSupervisor ? "Yes" : "No")}\n" +
                                             $"Is Manager: {(currentUser.IsManager ? "Yes" : "No")}\n\n" +
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