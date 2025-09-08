using ActiveDirectoryLookup.NetFramework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ActiveDirectoryLookup.Utilities
{
    /// <summary>
    /// Static interface class for Active Directory operations.
    /// Provides a simplified static API for common Active Directory lookup operations
    /// that can be easily integrated into shared libraries and static utility classes.
    /// 
    /// <para>
    /// <strong>Usage in Common/Static Libraries:</strong>
    /// </para>
    /// <para>
    /// This class is designed to be imported into common static classes where you need
    /// to provide Active Directory functionality without requiring instance management.
    /// </para>
    /// 
    /// <para>
    /// <strong>Dependencies:</strong>
    /// </para>
    /// <para>
    /// - ARC_AD_Lookup.dll (.NET 6 wrapper)
    /// </para>
    /// <para>
    /// - ActiveDirectoryLookup.NetFramework.dll (.NET Framework 4.8)
    /// </para>
    /// </summary>
    /// <example>
    /// <code>
    /// // In your common static class:
    /// public static class UserLookupUtilities
    /// {
    ///     public static List&lt;ActiveDirectoryUser&gt; FindUsers(string searchTerm)
    ///     {
    ///         return ActiveDirectoryHelper.SearchUsers(searchTerm);
    ///     }
    ///     
    ///     public static string GetUserEmail(string samAccountName)
    ///     {
    ///         var user = ActiveDirectoryHelper.GetUser(samAccountName);
    ///         return user?.Email ?? string.Empty;
    ///     }
    /// }
    /// </code>
    /// </example>
    public static class ActiveDirectoryHelper
    {
        /// <summary>
        /// Searches for Active Directory users by name using wildcard matching.
        /// This is a static wrapper around the ActiveDirectoryService.SearchUsersByName method.
        /// </summary>
        /// <param name="searchName">
        /// The name or partial name to search for. Case-insensitive search across
        /// displayName, common name (cn), and SAM account name fields.
        /// </param>
        /// <returns>
        /// A list of <see cref="ActiveDirectoryUser"/> objects representing the matching users.
        /// Returns an empty list if no users are found or if an error occurs.
        /// </returns>
        /// <remarks>
        /// This method creates a new ActiveDirectoryService instance for each call.
        /// For multiple operations, consider using the ActiveDirectoryService class directly
        /// to avoid the overhead of repeated instantiation.
        /// 
        /// Errors are handled internally and logged to System.Diagnostics.Debug.
        /// Returns an empty list on any failure to maintain static method simplicity.
        /// </remarks>
        /// <example>
        /// <code>
        /// var users = ActiveDirectoryHelper.SearchUsers("Smith");
        /// foreach (var user in users)
        /// {
        ///     Console.WriteLine($"{user.Name} - {user.Email}");
        /// }
        /// </code>
        /// </example>
        public static List<ActiveDirectoryUser> SearchUsers(string searchName)
        {
            if (string.IsNullOrWhiteSpace(searchName))
                return new List<ActiveDirectoryUser>();

            try
            {
                var service = new ActiveDirectoryService();
                return service.SearchUsersByName(searchName);
            }
            catch (Exception ex)
            {
                // Log error but return empty list to maintain static method simplicity
                System.Diagnostics.Debug.WriteLine($"ActiveDirectoryHelper.SearchUsers failed: {ex.Message}");
                return new List<ActiveDirectoryUser>();
            }
        }

        /// <summary>
        /// Gets a specific Active Directory user by their exact SAM account name.
        /// This is a static wrapper around the ActiveDirectoryService.SearchBySamAccountName method.
        /// </summary>
        /// <param name="samAccountName">
        /// The exact SAM account name (Windows logon name) to search for.
        /// </param>
        /// <returns>
        /// An <see cref="ActiveDirectoryUser"/> object if the user is found;
        /// otherwise, <see langword="null"/>.
        /// </returns>
        /// <remarks>
        /// This method creates a new ActiveDirectoryService instance for each call.
        /// Errors are handled internally and logged to System.Diagnostics.Debug.
        /// Returns null on any failure to maintain static method simplicity.
        /// </remarks>
        /// <example>
        /// <code>
        /// var user = ActiveDirectoryHelper.GetUser("jsmith");
        /// if (user != null)
        /// {
        ///     Console.WriteLine($"Found: {user.Name} - {user.Email}");
        /// }
        /// else
        /// {
        ///     Console.WriteLine("User not found");
        /// }
        /// </code>
        /// </example>
        public static ActiveDirectoryUser GetUser(string samAccountName)
        {
            if (string.IsNullOrWhiteSpace(samAccountName))
                return null;

            try
            {
                var service = new ActiveDirectoryService();
                var users = service.SearchBySamAccountName(samAccountName);
                return users.FirstOrDefault();
            }
            catch (Exception ex)
            {
                // Log error but return null to maintain static method simplicity
                System.Diagnostics.Debug.WriteLine($"ActiveDirectoryHelper.GetUser failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets the Active Directory information for the current logged-in user.
        /// This is a static wrapper around the ActiveDirectoryService.GetCurrentUserInfo method.
        /// </summary>
        /// <returns>
        /// An <see cref="ActiveDirectoryUser"/> object representing the current user if found;
        /// otherwise, <see langword="null"/>.
        /// </returns>
        /// <remarks>
        /// This method creates a new ActiveDirectoryService instance for each call.
        /// Uses Environment.UserName to determine the current user's SAM account name.
        /// Errors are handled internally and logged to System.Diagnostics.Debug.
        /// Returns null on any failure to maintain static method simplicity.
        /// </remarks>
        /// <example>
        /// <code>
        /// var currentUser = ActiveDirectoryHelper.GetCurrentUser();
        /// if (currentUser != null)
        /// {
        ///     Console.WriteLine($"Current user: {currentUser.Name}");
        ///     Console.WriteLine($"Email: {currentUser.Email}");
        /// }
        /// </code>
        /// </example>
        public static ActiveDirectoryUser GetCurrentUser()
        {
            try
            {
                var service = new ActiveDirectoryService();
                return service.GetCurrentUserInfo();
            }
            catch (Exception ex)
            {
                // Log error but return null to maintain static method simplicity
                System.Diagnostics.Debug.WriteLine($"ActiveDirectoryHelper.GetCurrentUser failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Tests the connectivity and functionality of the Active Directory service.
        /// This is a static wrapper around the ActiveDirectoryService.TestConnection method.
        /// </summary>
        /// <returns>
        /// <see langword="true"/> if the Active Directory service is working properly;
        /// otherwise, <see langword="false"/>.
        /// </returns>
        /// <remarks>
        /// This method creates a new ActiveDirectoryService instance for the test.
        /// Detailed diagnostic information is logged to System.Diagnostics.Debug regardless of success/failure.
        /// Use this method to verify Active Directory functionality before performing operations.
        /// </remarks>
        /// <example>
        /// <code>
        /// if (ActiveDirectoryHelper.TestConnection())
        /// {
        ///     Console.WriteLine("AD service is working");
        ///     // Proceed with AD operations
        /// }
        /// else
        /// {
        ///     Console.WriteLine("AD service is not available");
        ///     // Handle offline scenario
        /// }
        /// </code>
        /// </example>
        public static bool TestConnection()
        {
            try
            {
                var service = new ActiveDirectoryService();
                var success = service.TestConnection(out string message);
                
                // Log the detailed test results
                if (success)
                {
                    System.Diagnostics.Debug.WriteLine($"ActiveDirectoryHelper.TestConnection succeeded: {message}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"ActiveDirectoryHelper.TestConnection failed: {message}");
                }
                
                return success;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ActiveDirectoryHelper.TestConnection exception: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets the email address for a specific user by their SAM account name.
        /// This is a convenience method that combines user lookup and email extraction.
        /// </summary>
        /// <param name="samAccountName">
        /// The SAM account name of the user whose email address is needed.
        /// </param>
        /// <returns>
        /// The user's email address if found; otherwise, an empty string.
        /// </returns>
        /// <remarks>
        /// This is a convenience method for scenarios where you only need the email address.
        /// Returns an empty string if the user is not found or has no email address.
        /// </remarks>
        /// <example>
        /// <code>
        /// string email = ActiveDirectoryHelper.GetUserEmail("jsmith");
        /// if (!string.IsNullOrEmpty(email))
        /// {
        ///     Console.WriteLine($"User email: {email}");
        /// }
        /// </code>
        /// </example>
        public static string GetUserEmail(string samAccountName)
        {
            var user = GetUser(samAccountName);
            return user?.Email ?? string.Empty;
        }

        /// <summary>
        /// Gets the employee ID for a specific user by their SAM account name.
        /// This is a convenience method that combines user lookup and employee ID extraction.
        /// </summary>
        /// <param name="samAccountName">
        /// The SAM account name of the user whose employee ID is needed.
        /// </param>
        /// <returns>
        /// The user's employee ID if found; otherwise, an empty string.
        /// </returns>
        /// <remarks>
        /// This is a convenience method for scenarios where you only need the employee ID.
        /// Returns an empty string if the user is not found or has no employee ID.
        /// </remarks>
        /// <example>
        /// <code>
        /// string empId = ActiveDirectoryHelper.GetUserEmployeeId("jsmith");
        /// if (!string.IsNullOrEmpty(empId))
        /// {
        ///     Console.WriteLine($"Employee ID: {empId}");
        /// }
        /// </code>
        /// </example>
        public static string GetUserEmployeeId(string samAccountName)
        {
            var user = GetUser(samAccountName);
            return user?.EmployeeID ?? string.Empty;
        }

        /// <summary>
        /// Gets the display name for a specific user by their SAM account name.
        /// This is a convenience method that combines user lookup and display name extraction.
        /// </summary>
        /// <param name="samAccountName">
        /// The SAM account name of the user whose display name is needed.
        /// </param>
        /// <returns>
        /// The user's display name if found; otherwise, an empty string.
        /// </returns>
        /// <remarks>
        /// This is a convenience method for scenarios where you only need the display name.
        /// Returns an empty string if the user is not found or has no display name.
        /// </remarks>
        /// <example>
        /// <code>
        /// string displayName = ActiveDirectoryHelper.GetUserDisplayName("jsmith");
        /// if (!string.IsNullOrEmpty(displayName))
        /// {
        ///     Console.WriteLine($"User name: {displayName}");
        /// }
        /// </code>
        /// </example>
        public static string GetUserDisplayName(string samAccountName)
        {
            var user = GetUser(samAccountName);
            return user?.Name ?? string.Empty;
        }
    }
}