using ActiveDirectoryLookup.NetFramework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ActiveDirectoryLookup.Utilities
{
    /// <summary>
    /// Represents an Active Directory user with essential properties for .NET 6 applications.
    /// This is a wrapper class that corresponds to the .NET Framework ActiveDirectoryUser class.
    /// </summary>
    /// <remarks>
    /// This class provides a .NET 6 compatible interface to Active Directory user data
    /// while delegating the actual AD operations to a .NET Framework 4.8 DLL.
    /// </remarks>
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
    /// Provides Active Directory search and lookup functionality for .NET 6 applications.
    /// This service acts as a wrapper around the .NET Framework 4.8 ActiveDirectoryLookup.NetFramework DLL
    /// to provide Active Directory access in environments where System.DirectoryServices is not supported.
    /// 
    /// <para>
    /// <strong>Architecture:</strong>
    /// </para>
    /// <para>
    /// This wrapper pattern allows .NET 6 applications to use Active Directory functionality
    /// by delegating to a .NET Framework 4.8 DLL that has full System.DirectoryServices support.
    /// </para>
    /// 
    /// <para>
    /// <strong>Usage for Client Applications:</strong>
    /// </para>
    /// <para>
    /// 1. Add reference to ARC_AD_Lookup.dll (.NET 6 wrapper)
    /// </para>
    /// <para>
    /// 2. Ensure ActiveDirectoryLookup.NetFramework.dll (.NET Framework 4.8) is available
    /// </para>
    /// <para>
    /// 3. Application must run with appropriate Active Directory permissions
    /// </para>
    /// <para>
    /// 4. All methods throw exceptions on failure - use try/catch for error handling
    /// </para>
    /// </summary>
    /// <example>
    /// <code>
    /// // Basic usage in .NET 6 application
    /// var adService = new ActiveDirectoryService();
    /// 
    /// try
    /// {
    ///     var users = adService.SearchUsersByName("Smith");
    ///     foreach (var user in users)
    ///     {
    ///         Console.WriteLine($"{user.Name} - {user.Email} - {user.EmployeeID}");
    ///     }
    /// }
    /// catch (Exception ex)
    /// {
    ///     Console.WriteLine($"AD Search failed: {ex.Message}");
    /// }
    /// </code>
    /// </example>
    public class ActiveDirectoryService
    {
        private readonly ActiveDirectoryLookup.NetFramework.ActiveDirectoryService _frameworkService;

        /// <summary>
        /// Initializes a new instance of the ActiveDirectoryService.
        /// Creates an instance of the underlying .NET Framework 4.8 service.
        /// </summary>
        /// <exception cref="System.IO.FileNotFoundException">
        /// Thrown when the ActiveDirectoryLookup.NetFramework.dll cannot be found.
        /// </exception>
        /// <exception cref="System.TypeLoadException">
        /// Thrown when the .NET Framework service cannot be instantiated.
        /// </exception>
        public ActiveDirectoryService()
        {
            _frameworkService = new ActiveDirectoryLookup.NetFramework.ActiveDirectoryService();
        }

        /// <summary>
        /// Searches for Active Directory users by name using wildcard matching.
        /// Searches across displayName, common name (cn), and SAM account name fields.
        /// </summary>
        /// <param name="searchName">
        /// The name or partial name to search for. Wildcard characters are automatically added.
        /// Case-insensitive search across multiple name fields.
        /// </param>
        /// <returns>
        /// A list of <see cref="ActiveDirectoryUser"/> objects representing the matching users.
        /// Limited to 100 results to prevent overwhelming the application.
        /// Returns an empty list if no users are found.
        /// </returns>
        /// <exception cref="System.Exception">
        /// Thrown when the Active Directory search fails. The exception message contains
        /// diagnostic information including the framework DLL location and environment details.
        /// </exception>
        /// <remarks>
        /// This method uses the current user's Windows credentials for Active Directory authentication.
        /// Requires appropriate read permissions on the Active Directory domain.
        /// The search is performed against the default domain context.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// 
        /// try
        /// {
        ///     var users = service.SearchUsersByName("John");
        ///     Console.WriteLine($"Found {users.Count} users");
        ///     
        ///     foreach (var user in users)
        ///     {
        ///         Console.WriteLine($"{user.Name} ({user.SamAccountName})");
        ///         Console.WriteLine($"  Email: {user.Email}");
        ///         Console.WriteLine($"  Employee ID: {user.EmployeeID}");
        ///     }
        /// }
        /// catch (Exception ex)
        /// {
        ///     Console.WriteLine($"Search failed: {ex.Message}");
        /// }
        /// </code>
        /// </example>
        public List<ActiveDirectoryUser> SearchUsersByName(string searchName)
        {
            var users = new List<ActiveDirectoryUser>();

            try
            {
                var result = _frameworkService.SearchUsersByName(searchName);

                if (!result.Success)
                {
                    throw new Exception(result.ErrorMessage);
                }

                // Convert from .NET Framework objects to .NET 6 objects
                foreach (var frameworkUser in result.Users)
                {
                    users.Add(new ActiveDirectoryUser
                    {
                        EmployeeID = frameworkUser.EmployeeID,
                        Email = frameworkUser.Email,
                        Name = frameworkUser.Name,
                        SamAccountName = frameworkUser.SamAccountName
                    });
                }
            }
            catch (Exception ex)
            {
                // Add more diagnostic information
                var diagnosticInfo = $"Framework DLL Location: {typeof(ActiveDirectoryLookup.NetFramework.ActiveDirectoryService).Assembly.Location}\n" +
                                   $"Current User: {Environment.UserName}\n" +
                                   $"Domain: {Environment.UserDomainName}\n" +
                                   $"Machine: {Environment.MachineName}";
                
                throw new Exception($"Error searching Active Directory: {ex.Message}\n\nDiagnostic Info:\n{diagnosticInfo}", ex);
            }

            return users;
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
        /// A list of <see cref="ActiveDirectoryUser"/> objects. Should contain 0 or 1 user
        /// since SAM account names are unique. Returns an empty list if user is not found.
        /// </returns>
        /// <exception cref="System.Exception">
        /// Thrown when the Active Directory search fails.
        /// </exception>
        /// <remarks>
        /// This method is ideal for looking up a specific user when you know their exact logon name.
        /// Uses the current user's Windows credentials for Active Directory authentication.
        /// Search is limited to 1 result since SAM account names should be unique.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// 
        /// try
        /// {
        ///     var users = service.SearchBySamAccountName("jsmith");
        ///     
        ///     if (users.Count > 0)
        ///     {
        ///         var user = users[0];
        ///         Console.WriteLine($"Found user: {user.Name}");
        ///         Console.WriteLine($"Email: {user.Email}");
        ///         Console.WriteLine($"Employee ID: {user.EmployeeID}");
        ///     }
        ///     else
        ///     {
        ///         Console.WriteLine("User not found");
        ///     }
        /// }
        /// catch (Exception ex)
        /// {
        ///     Console.WriteLine($"Search failed: {ex.Message}");
        /// }
        /// </code>
        /// </example>
        public List<ActiveDirectoryUser> SearchBySamAccountName(string samAccountName)
        {
            var users = new List<ActiveDirectoryUser>();

            try
            {
                var result = _frameworkService.SearchBySamAccountName(samAccountName);

                if (!result.Success)
                {
                    throw new Exception(result.ErrorMessage);
                }

                // Convert from .NET Framework objects to .NET 6 objects
                foreach (var frameworkUser in result.Users)
                {
                    users.Add(new ActiveDirectoryUser
                    {
                        EmployeeID = frameworkUser.EmployeeID,
                        Email = frameworkUser.Email,
                        Name = frameworkUser.Name,
                        SamAccountName = frameworkUser.SamAccountName
                    });
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error searching Active Directory by SAM account: {ex.Message}", ex);
            }

            return users;
        }

        /// <summary>
        /// Tests the connectivity and functionality of the Active Directory service.
        /// Performs a comprehensive test including connection verification and current user lookup.
        /// </summary>
        /// <param name="message">
        /// When this method returns, contains detailed diagnostic information about the test results.
        /// Includes framework DLL location, environment details, and current user information if found.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the connection test passed and AD functionality is working;
        /// otherwise, <see langword="false"/>.
        /// </returns>
        /// <remarks>
        /// This method performs the most comprehensive test available by:
        /// - Testing basic Active Directory connectivity
        /// - Verifying user search functionality and permissions
        /// - Attempting to find the current user's profile in AD
        /// - Providing detailed diagnostic information for troubleshooting
        /// 
        /// Use this method for initial setup validation and troubleshooting scenarios.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// 
        /// if (service.TestConnection(out string message))
        /// {
        ///     Console.WriteLine("AD Service is working properly");
        ///     Console.WriteLine(message); // Shows detailed success info
        /// }
        /// else
        /// {
        ///     Console.WriteLine("AD Service test failed");
        ///     Console.WriteLine(message); // Shows error details and diagnostics
        /// }
        /// </code>
        /// </example>
        public bool TestConnection(out string message)
        {
            try
            {
                var result = _frameworkService.TestConnectionAndSearchCurrentUser();
                
                // Add diagnostic information to the message
                var diagnosticInfo = $"Framework DLL: {typeof(ActiveDirectoryLookup.NetFramework.ActiveDirectoryService).Assembly.Location}\n" +
                                   $"User: {Environment.UserName}\n" +
                                   $"Domain: {Environment.UserDomainName}\n" +
                                   $"Machine: {Environment.MachineName}\n\n";
                
                message = diagnosticInfo + result.ErrorMessage;
                return result.Success;
            }
            catch (Exception ex)
            {
                message = $"Framework DLL Error: {ex.Message}\n" +
                         $"Assembly: {typeof(ActiveDirectoryLookup.NetFramework.ActiveDirectoryService).Assembly.Location}";
                return false;
            }
        }

        /// <summary>
        /// Gets the Active Directory information for the current logged-in user.
        /// This is a convenience method that searches for the current user's SAM account name.
        /// </summary>
        /// <returns>
        /// An <see cref="ActiveDirectoryUser"/> object representing the current user if found;
        /// otherwise, <see langword="null"/> if the current user is not found in Active Directory.
        /// </returns>
        /// <exception cref="System.Exception">
        /// Thrown when the Active Directory search fails.
        /// </exception>
        /// <remarks>
        /// This method uses <see cref="Environment.UserName"/> to get the current user's SAM account name
        /// and then performs a search using <see cref="SearchBySamAccountName"/>.
        /// </remarks>
        /// <example>
        /// <code>
        /// var service = new ActiveDirectoryService();
        /// 
        /// try
        /// {
        ///     var currentUser = service.GetCurrentUserInfo();
        ///     
        ///     if (currentUser != null)
        ///     {
        ///         Console.WriteLine($"Current user: {currentUser.Name}");
        ///         Console.WriteLine($"Email: {currentUser.Email}");
        ///         Console.WriteLine($"Employee ID: {currentUser.EmployeeID}");
        ///     }
        ///     else
        ///     {
        ///         Console.WriteLine("Current user not found in AD");
        ///     }
        /// }
        /// catch (Exception ex)
        /// {
        ///     Console.WriteLine($"Failed to get current user info: {ex.Message}");
        /// }
        /// </code>
        /// </example>
        public ActiveDirectoryUser GetCurrentUserInfo()
        {
            string samAccountName = Environment.UserName;
            var users = SearchBySamAccountName(samAccountName);
            return users.FirstOrDefault();
        }
    }
}