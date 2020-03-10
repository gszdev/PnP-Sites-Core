﻿using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;
using System.Net;
#if !NETSTANDARD2_0
using System.Data.SqlClient;
#endif
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Security.Cryptography.X509Certificates;
using OfficeDevPnP.Core.Utilities;
using Newtonsoft.Json.Linq;

namespace OfficeDevPnP.Core.Tests
{
    static class TestCommon
    {
#if NETSTANDARD2_0
        private static Configuration configuration = null;
#endif

        public static string AppSetting(string key)
        {
#if !NETSTANDARD2_0
            return ConfigurationManager.AppSettings[key];
#else
            try
            {
                return configuration.AppSettings.Settings[key].Value;
            }
            catch
            {
                return null;
            }
#endif
        }

        #region Constructor
        static TestCommon()
        {
#if NETSTANDARD2_0
            // Load configuration in a way that's compatible with a .Net Core test project as well
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap
            {
                ExeConfigFilename = @"..\..\App.config" //Path to your config file
            };
            configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
#endif

            // Read configuration data
            TenantUrl = AppSetting("SPOTenantUrl");
            DevSiteUrl = AppSetting("SPODevSiteUrl");
            DefaultSiteOwner = AppSetting("DefaultSiteOwner");
            if (string.IsNullOrEmpty(DefaultSiteOwner))
            {
#if !ONPREMISES
                DefaultSiteOwner = AppSetting("SPOUserName");
#else                    
                DefaultSiteOwner = $"{AppSetting("OnPremDomain")}\\{AppSetting("OnPremUserName")}";
#endif
            }

#if !SP2013 && !SP2016
            if (string.IsNullOrEmpty(TenantUrl))
            {
                throw new ConfigurationErrorsException("Tenant site Url in App.config are not set up.");
            }
#endif
            if (string.IsNullOrEmpty(DevSiteUrl))
            {
                throw new ConfigurationErrorsException("Dev site url in App.config are not set up.");
            }



            // Trim trailing slashes
            TenantUrl = TenantUrl.TrimEnd(new[] { '/' });
            DevSiteUrl = DevSiteUrl.TrimEnd(new[] { '/' });

            var spoCredentialManagerLabel = AppSetting("SPOCredentialManagerLabel");

            if (!string.IsNullOrEmpty(spoCredentialManagerLabel))
            {
                var tempCred = Core.Utilities.CredentialManager.GetCredential(spoCredentialManagerLabel);

                UserName = tempCred.UserName;
                Password = tempCred.SecurePassword;

                // username in format domain\user means we're testing in on-premises
                if (tempCred.UserName.IndexOf("\\") > 0)
                {
                    string[] userParts = tempCred.UserName.Split('\\');
                    Credentials = new NetworkCredential(userParts[1], tempCred.SecurePassword, userParts[0]);
                }
                else
                {
                    Credentials = new SharePointOnlineCredentials(tempCred.UserName, tempCred.SecurePassword);
                }
            }
            else
            {
#if ONPREMISES
                var onPremUseHighTrustCertificateValue = AppSetting("OnPremUseHighTrustCertificate");

                if (!string.IsNullOrEmpty(onPremUseHighTrustCertificateValue))
                {
                    bool tempOnPremUseHighTrustCertificate = false;
                    if (bool.TryParse(onPremUseHighTrustCertificateValue, out tempOnPremUseHighTrustCertificate))
                    {
                        OnPremUseHighTrustCertificate = tempOnPremUseHighTrustCertificate;
                    }
                }

                HighTrustClientId = AppSetting("HighTrustClientId");
                if (string.IsNullOrEmpty(HighTrustClientId))
                {
                    HighTrustClientId = AppSetting("AppId");
                }

                HighTrustCertificatePassword = AppSetting("HighTrustCertificatePassword");
                HighTrustCertificatePath = AppSetting("HighTrustCertificatePath");
                HighTrustIssuerId = AppSetting("HighTrustIssuerId");
                HighTrustBehalfOfUserLoginName = AppSetting("HighTrustBehalfOfUserLoginName");

                if (!String.IsNullOrEmpty(AppSetting("HighTrustCertificateStoreName")))
                {
                    StoreName result;
                    if (Enum.TryParse(AppSetting("HighTrustCertificateStoreName"), out result))
                    {
                        HighTrustCertificateStoreName = result;
                    }
                }

                if (!String.IsNullOrEmpty(AppSetting("HighTrustCertificateStoreLocation")))
                {
                    StoreLocation result;
                    if (Enum.TryParse(AppSetting("HighTrustCertificateStoreLocation"), out result))
                    {
                        HighTrustCertificateStoreLocation = result;
                    }
                }

                HighTrustCertificateStoreThumbprint = AppSetting("HighTrustCertificateStoreThumbprint").Replace(" ", string.Empty);
#endif
                var spoUseAppOnlyValue = AppSetting("SPOUseAppOnly");
                if (!string.IsNullOrEmpty(spoUseAppOnlyValue))
                {
                    var tempSPOUseAppOnly = false;
                    if (bool.TryParse(spoUseAppOnlyValue, out tempSPOUseAppOnly))
                    {
                        SPOUseAppOnly = tempSPOUseAppOnly;
                    }
                }

#if !ONPREMISES
                if (!String.IsNullOrEmpty(AppSetting("SPOUserName")) &&
                    !String.IsNullOrEmpty(AppSetting("SPOPassword")))
                {
                    UserName = AppSetting("SPOUserName");
                    var password = AppSetting("SPOPassword");

                    Password = EncryptionUtility.ToSecureString(password);
                    Credentials = new SharePointOnlineCredentials(UserName, Password);
                }
                else
#endif
                if (OnPremUseHighTrustCertificate == false
                    && !String.IsNullOrEmpty(AppSetting("OnPremUserName"))
                    && !String.IsNullOrEmpty(AppSetting("OnPremDomain"))
                    && !String.IsNullOrEmpty(AppSetting("OnPremPassword"))
                    )
                {
                    Password = EncryptionUtility.ToSecureString(AppSetting("OnPremPassword"));
                    Credentials = new NetworkCredential(AppSetting("OnPremUserName"), Password, AppSetting("OnPremDomain"));
                }
                else if (OnPremUseHighTrustCertificate == false
                    && !String.IsNullOrEmpty(AppSetting("AppId"))
                    && !String.IsNullOrEmpty(AppSetting("AppSecret"))
                    )
                {
                    AppId = AppSetting("AppId");
                    AppSecret = AppSetting("AppSecret");

                    IsAppOnlyTesting = true;
                }
                else if (OnPremUseHighTrustCertificate == true
                    || (
                        !String.IsNullOrEmpty(AppSetting("AppId"))
                        && !String.IsNullOrEmpty(AppSetting("HighTrustIssuerId"))
                        )
                    )
                {
                    // AppId = AppSetting("AppId");
                    HighTrustClientId = AppSetting("HighTrustClientId");
                    if (string.IsNullOrEmpty(HighTrustClientId))
                    {
                        HighTrustClientId = AppSetting("AppId");
                    }
                    HighTrustCertificatePassword = AppSetting("HighTrustCertificatePassword");
                    HighTrustCertificatePath = AppSetting("HighTrustCertificatePath");
                    HighTrustIssuerId = AppSetting("HighTrustIssuerId");
                    HighTrustBehalfOfUserLoginName = AppSetting("HighTrustBehalfOfUserLoginName");

                    if (!String.IsNullOrEmpty(AppSetting("HighTrustCertificateStoreName")))
                    {
                        StoreName result;
                        if (Enum.TryParse(AppSetting("HighTrustCertificateStoreName"), out result))
                        {
                            HighTrustCertificateStoreName = result;
                        }
                    }
                    if (!String.IsNullOrEmpty(AppSetting("HighTrustCertificateStoreLocation")))
                    {
                        StoreLocation result;
                        if (Enum.TryParse(AppSetting("HighTrustCertificateStoreLocation"), out result))
                        {
                            HighTrustCertificateStoreLocation = result;
                        }
                    }
                    HighTrustCertificateStoreThumbprint = AppSetting("HighTrustCertificateStoreThumbprint").Replace(" ", string.Empty);

                    if (string.IsNullOrEmpty(HighTrustBehalfOfUserLoginName))
                    {
                        IsAppOnlyTesting = true;
                    }
                }
                else
                {
                    throw new ConfigurationErrorsException("Tenant credentials in App.config are not set up.");
                }
            }
        }
#endregion

#region Properties
        public static string TenantUrl { get; set; }
        public static string DevSiteUrl { get; set; }
        static string UserName { get; set; }
        static SecureString Password { get; set; }
        public static ICredentials Credentials { get; set; }
        public static string AppId { get; set; }
        static string AppSecret { get; set; }
        public static string DefaultSiteOwner { get; set; }

        /// <summary>
        ///
        /// </summary>
        public static bool IsAppOnlyTesting { get; private set; }

        public static String AzureADCertificateFilePath
        {
            get
            {
                return AppSetting("AzureADCertificateFilePath");
            }
        }

        /// <summary>
        /// The ClientID is the web application's client ID (GUID) that was generated on appregnew.aspx / same value as the HighTrustIssuerId
        /// <remarks>
        /// https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/package-and-publish-high-trust-sharepoint-add-ins
        ///  
        /// If the high-trust SharePoint Add-in has its own certificate that it is not sharing with other SharePoint Add-ins, the IssuerId is the same as the ClientId 
        /// </remarks> 
        /// </summary> 
        public static string HighTrustClientId { get; set; }

        public static bool OnPremUseHighTrustCertificate { get; set; }

        public static string HighTrustBehalfOfUserLoginName { get; set; }

        public static bool SPOUseAppOnly { get; set; }

        /// <summary>
        /// The path to the PFX file for the High Trust
        /// </summary>
        public static String HighTrustCertificatePath { get; set; }

        /// <summary>
        /// The password of the PFX file for the High Trust
        /// </summary>
        public static String HighTrustCertificatePassword { get; set; }

        /// <summary>
        /// The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer
        /// </summary>
        public static String HighTrustIssuerId { get; set; }

        /// <summary>
        /// The name of the store in the Windows certificate store where the High Trust certificate is stored
        /// </summary>
        public static StoreName? HighTrustCertificateStoreName { get; set; }

        /// <summary>
        /// The location of the High Trust certificate in the Windows certificate store
        /// </summary>
        public static StoreLocation? HighTrustCertificateStoreLocation { get; set; }

        /// <summary>
        /// The thumbprint / hash of the High Trust certificate in the Windows certificate store
        /// </summary>
        public static string HighTrustCertificateStoreThumbprint { get; set; }

        public static string TestWebhookUrl
        {
            get
            {
                return AppSetting("WebHookTestUrl");
            }
        }

        public static String AzureStorageKey
        {
            get
            {
                return AppSetting("AzureStorageKey");
            }
        }
        public static String TestAutomationDatabaseConnectionString
        {
            get
            {
                return AppSetting("TestAutomationDatabaseConnectionString");
            }
        }
        public static String AzureADCertPfxPassword
        {
            get
            {
                return AppSetting("AzureADCertPfxPassword");
            }
        }
        public static String AzureADClientId
        {
            get
            {
                return AppSetting("AzureADClientId");
            }
        }
        public static String NoScriptSite
        {
            get
            {
                return AppSetting("NoScriptSite");
            }
        }
        public static String ScriptSite
        {
            get
            {
                return AppSetting("ScriptSite");
            }
        }
#endregion

#region Methods
        public static ClientContext CreateClientContext()
        {
            return CreateContext(DevSiteUrl, Credentials);
        }

        public static ClientContext CreateClientContext(string url)
        {
            return CreateContext(url, Credentials);
        }

        public static ClientContext CreateTenantClientContext()
        {
            return CreateContext(TenantUrl, Credentials);
        }

        public static PnPClientContext CreatePnPClientContext(int retryCount = 10, int delay = 500)
        {
            PnPClientContext context;
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                AuthenticationManager am = new AuthenticationManager();
                ClientContext clientContext = null;

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    //clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, Core.Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)), AppId, AppSecret, acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login");
                    clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret, AzureEnvironment.PPE);
                }
                else if (new Uri(DevSiteUrl).DnsSafeHost.Contains("sharepoint.de"))
                {
                    clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret, AzureEnvironment.Germany);
                }
                else
                {
                    clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret);
                }
                context = PnPClientContext.ConvertFrom(clientContext, retryCount, delay);
            }
            else
            {
                context = new PnPClientContext(DevSiteUrl, retryCount, delay);
                context.Credentials = Credentials;
            }

            context.RequestTimeout = 1000 * 60 * 15;
            return context;
        }


        public static bool AppOnlyTesting()
        {
            /*
            if (!String.IsNullOrEmpty(AppSetting("AppId")) &&
                !String.IsNullOrEmpty(AppSetting("AppSecret")) &&
                String.IsNullOrEmpty(AppSetting("SPOCredentialManagerLabel")) &&
                String.IsNullOrEmpty(AppSetting("SPOUserName")) &&
                String.IsNullOrEmpty(AppSetting("SPOPassword")) &&
                String.IsNullOrEmpty(AppSetting("OnPremUserName")) &&
                String.IsNullOrEmpty(AppSetting("OnPremDomain")) &&
                String.IsNullOrEmpty(AppSetting("OnPremPassword")))
            {
                return true;
            }
            else
            {
                return false;
            }           
             */
            return IsAppOnlyTesting;
        }

#if !NETSTANDARD2_0
        public static bool TestAutomationSQLDatabaseAvailable()
        {
            string connectionString = TestAutomationDatabaseConnectionString;
            if (!String.IsNullOrEmpty(connectionString))
            {
                try
                {
                    var con = new SqlConnectionStringBuilder(connectionString);
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        return (conn.State == ConnectionState.Open);
                    }
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }
#endif

        private static ClientContext CreateContext(string contextUrl, ICredentials credentials)
        {
            ClientContext context =  null;
            if (!String.IsNullOrEmpty(AppId)
                && !String.IsNullOrEmpty(AppSecret))
            {
                AuthenticationManager am = new AuthenticationManager();

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
#if !NETSTANDARD2_0
                    context = am.GetAppOnlyAuthenticatedContext(
                        contextUrl,
                        Core.Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)),
                        AppId,
                        AppSecret,
                        acsHostUrl: "windows-ppe.net",
                        globalEndPointPrefix: "login");
#endif
                }
                else
                {
                    context = am.GetAppOnlyAuthenticatedContext(contextUrl, AppId, AppSecret);
                }
            }
#if ONPREMISES
            else if (OnPremUseHighTrustCertificate
                && ( 
                    ! string.IsNullOrEmpty(HighTrustClientId)  
                    && !string.IsNullOrEmpty(HighTrustCertificatePath) 
                    && !string.IsNullOrEmpty(HighTrustCertificatePassword) 
                    && !string.IsNullOrEmpty(HighTrustIssuerId) 
                )) 
            { 
                // Instantiate a ClientContext object based on the defined high trust certificate 
                if (!string.IsNullOrEmpty(HighTrustBehalfOfUserLoginName)) 
                { 
                    context = new AuthenticationManager().GetHighTrustCertificateAppAuthenticatedContext( 
                        contextUrl, 
                        HighTrustClientId, 
                        HighTrustCertificatePath, 
                        HighTrustCertificatePassword, 
                        HighTrustIssuerId, 
                        HighTrustBehalfOfUserLoginName);                         
                } 
                else 
                { 
                    context = new AuthenticationManager().GetHighTrustCertificateAppOnlyAuthenticatedContext( 
                        contextUrl, 
                        HighTrustClientId, 
                        HighTrustCertificatePath, 
                        HighTrustCertificatePassword, 
                        HighTrustIssuerId); 
                }                 
            } 
            else if (OnPremUseHighTrustCertificate 
                &&  
                    ( 
                        string.IsNullOrEmpty(HighTrustClientId) 
                        && !!HighTrustCertificateStoreName.HasValue 
                        && !!HighTrustCertificateStoreLocation.HasValue 
                        && !string.IsNullOrEmpty(HighTrustCertificateStoreThumbprint) 
                        && !string.IsNullOrEmpty(HighTrustIssuerId) 
                    ) 
                ) 
            { 
                // Instantiate a ClientContext object based on the defined high trust certificate 
                if (!string.IsNullOrEmpty(HighTrustBehalfOfUserLoginName)) 
                { 
                    context = new AuthenticationManager().GetHighTrustCertificateAppAuthenticatedContext( 
                        contextUrl, 
                        HighTrustClientId, 
                        HighTrustCertificateStoreName.Value, 
                        HighTrustCertificateStoreLocation.Value, 
                        HighTrustCertificateStoreThumbprint, 
                        HighTrustIssuerId, 
                        HighTrustBehalfOfUserLoginName); 
                } 
                else 
                { 
                    context = new AuthenticationManager().GetHighTrustCertificateAppOnlyAuthenticatedContext( 
                        contextUrl, 
                        HighTrustClientId, 
                        HighTrustCertificateStoreName.Value, 
                        HighTrustCertificateStoreLocation.Value, 
                        HighTrustCertificateStoreThumbprint, 
                        HighTrustIssuerId); 
                } 
            } 
#endif
            else
            {
                context = new ClientContext(contextUrl);
                context.Credentials = credentials;
            }

            context.RequestTimeout = 1000 * 60 * 15;
            return context;
        }

#endregion


#if !ONPREMISES
        public static string AcquireTokenAsync(string resource, string scope = null)
        {
            var tenantId = TenantExtensions.GetTenantIdByUrl(TestCommon.AppSetting("SPOTenantUrl"));
            //var tenantId = GetTenantIdByUrl(TestCommon.AppSetting("SPOTenantUrl"));
            if (tenantId == null) return null;

            var clientId = TestCommon.AppSetting("AppId");
            var clientSecret = TestCommon.AppSetting("AppSecret");
            var username = UserName;
            var password = EncryptionUtility.ToInsecureString(Password);

            string body;
            string response;
            if (scope == null) // use v1 endpoint
            {
                body = $"grant_type=password&client_id={clientId}&client_secret={clientSecret}&username={username}&password={password}&resource={resource}";
                response = HttpHelper.MakePostRequestForString($"https://login.microsoftonline.com/{tenantId}/oauth2/token", body, "application/x-www-form-urlencoded");
            }
            else // use v2 endpoint
            {
                body = $"grant_type=password&client_id={clientId}&username={username}&password={password}&scope={scope}";
                response = HttpHelper.MakePostRequestForString($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", body, "application/x-www-form-urlencoded");
            }

            var json = JToken.Parse(response);
            return json["access_token"].ToString();
        }
#endif
        private static Assembly _newtonsoftAssembly;
        private static string _assemblyName;

        public static void FixAssemblyResolving(string assemblyName)
        {
            _assemblyName = assemblyName;
            _newtonsoftAssembly = Assembly.LoadFrom(Path.Combine(AssemblyDirectory, $"{assemblyName}.dll"));
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private static string AssemblyDirectory
        {
            get
            {
                var codeBase = Assembly.GetExecutingAssembly().CodeBase;
                var uri = new UriBuilder(codeBase);
                var path = Uri.UnescapeDataString(uri.Path);

                return Path.GetDirectoryName(path);
            }
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return args.Name.StartsWith(_assemblyName) ? _newtonsoftAssembly : AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(assembly => assembly.FullName == args.Name);
        }

        public static void DeleteFile(ClientContext ctx, string serverRelativeFileUrl)
        {
            var file = ctx.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
            ctx.Load(file, f => f.Exists);
            ctx.ExecuteQueryRetry();

            if (file.Exists)
            {
                file.DeleteObject();
                ctx.ExecuteQueryRetry();
            }
        }
    }
}
