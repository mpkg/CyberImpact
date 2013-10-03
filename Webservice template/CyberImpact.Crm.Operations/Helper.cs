using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;

namespace CyberImpact.Crm.Operations
{
    /// <summary>
    /// Class for operations with CRM system
    /// </summary>
    public static class Helper
    {
        /// <summary>
        /// Check SSL certificate handler
        /// </summary>
        private static bool AcceptAllCertificatePolicy(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        /// <summary>
        /// Creates client to CRM 2011 web service using configured URL
        /// </summary>
        public static IOrganizationService GetCrmConnection()
            { return GetCrmConnection(null); }

        /// <summary>
        /// Creates client to CRM 2011 web service with specified URL
        /// </summary>
        public static IOrganizationService GetCrmConnection(Uri crmUrl)
        {
            if (crmUrl == null)
                crmUrl = new Uri(Properties.Settings.Default.CrmUrl);

            if (Properties.Settings.Default.UseWindowsAuthentication)
            {
                //to ignore certificates errors
                ServicePointManager.ServerCertificateValidationCallback = AcceptAllCertificatePolicy;
                IServiceManagement<IOrganizationService> orgServiceManagement =
                   ServiceConfigurationFactory.CreateManagement<IOrganizationService>(crmUrl);
                AuthenticationCredentials credentials = new AuthenticationCredentials();
                credentials.ClientCredentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
                return new OrganizationServiceProxy(orgServiceManagement, credentials.ClientCredentials);
            }
            else
            {
                ClientCredentials Credentials = new ClientCredentials();
                ClientCredentials deviceCredentials = new ClientCredentials();
                Credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
                deviceCredentials = Microsoft.Crm.Services.Utility.DeviceIdManager.LoadDeviceCredentials();

                //to ignore certificates errors
                ServicePointManager.ServerCertificateValidationCallback = AcceptAllCertificatePolicy;

                Uri OrganizationUri = crmUrl;
                Uri HomeRealmUri = null;
                OrganizationServiceProxy _serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials,
                    deviceCredentials);
                _serviceProxy.ClientCredentials.UserName.UserName = Properties.Settings.Default.UserName;
                _serviceProxy.ClientCredentials.UserName.Password = Properties.Settings.Default.UserPassword;
                _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
                return (IOrganizationService)_serviceProxy;
            }
            
        }

        /// <summary>
        /// Creates client to CRM 2011 web service with specified credentials using configured URL
        /// </summary>
        public static IOrganizationService GetCrmConnection(string username, string password)
            { return GetCrmConnection(username, password, null); }

        /// <summary>
        /// Creates client to CRM 2011 web service with specified credentials with specified URL
        /// </summary>
        public static IOrganizationService GetCrmConnection(string username, string password, Uri crmUrl)
        {
            if (crmUrl == null)
                crmUrl = new Uri(Properties.Settings.Default.CrmUrl);

            if (Properties.Settings.Default.UseWindowsAuthentication)
            {
                //to ignore certificates errors
                ServicePointManager.ServerCertificateValidationCallback = AcceptAllCertificatePolicy;
                IServiceManagement<IOrganizationService> orgServiceManagement =
                   ServiceConfigurationFactory.CreateManagement<IOrganizationService>(crmUrl);
                AuthenticationCredentials credentials = new AuthenticationCredentials();
                credentials.ClientCredentials.Windows.ClientCredential = new System.Net.NetworkCredential(username, password);
                return new OrganizationServiceProxy(orgServiceManagement, credentials.ClientCredentials);
            }
            else
            {
                ClientCredentials Credentials = new ClientCredentials();
                ClientCredentials deviceCredentials = new ClientCredentials();
                Credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
                deviceCredentials = Microsoft.Crm.Services.Utility.DeviceIdManager.LoadDeviceCredentials();

                //to ignore certificates errors
                ServicePointManager.ServerCertificateValidationCallback = AcceptAllCertificatePolicy;

                Uri OrganizationUri = crmUrl;
                Uri HomeRealmUri = null;
                OrganizationServiceProxy _serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials,
                    deviceCredentials);
                _serviceProxy.ClientCredentials.UserName.UserName = username;
                _serviceProxy.ClientCredentials.UserName.Password = password;
                _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
                return (IOrganizationService)_serviceProxy;
            }
        }
    }

    public class Customer
    {
        public DateTime dateGabarit { get; set; }
        public int noDuClient { get; set; }
        public string nomDuCompte { get; set; }
        public string telephone { get; set; }
        public string adresse1 { get; set; }
        public string adresse2 { get; set; }
        public string ville { get; set; }
        public string province { get; set; }
        public string pays { get; set; }
        public string codePostal { get; set; }
        public DateTime dateDeEssai { get; set; }
        public int nombreDeLogin { get; set; }
        public bool importationCompletee { get; set; }
        public DateTime dateImportation { get; set; }
        public bool gabaritCree { get; set; }
        public bool envoiComplete { get; set; }
        public DateTime dateEnvoi { get; set; }
        public DateTime dateActivation { get; set; }
        public int forfait { get; set; }
        public bool osbl { get; set; }
        public bool premium { get; set; }
        public bool actif { get; set; }
        public bool envoiPermis { get; set; }
        public bool revendeur { get; set; }
        public string noDeRevendeur { get; set; }
        public Money revenuForfait { get; set; }
        public Money revenuExtra { get; set; }
        public string courrielFrom { get; set; }
        public DateTime expirationDemo { get; set; }
        public bool bDSupprimee { get; set; }
        public bool accesAPI { get; set; }
        public bool accesAPISansOptIn { get; set; }
        public string entrepriseFacturation { get; set; }
        public int methodeDePaiement { get; set; }
    }

    public class Contact
    {
        public int noDuClient { get; set; }
        public string nom { get; set; }
        public string prenom { get; set; }
        public bool contactTechnique { get; set; }
        public bool contactFacturation { get; set; }
        public bool cyberBulletin { get; set; }
        public int langue { get; set; }
        public bool actif { get; set; }
    }
}
