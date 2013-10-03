using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml;
using CyberImpact.Crm.Operations;
using Microsoft.Xrm.Sdk;
using System.Xml.Serialization;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk.Client;

namespace CyberImpact.Crm.WebServices
{
    /// <summary>
    /// TODO
    /// </summary>
    [WebService(Namespace = "http://cyberimpact.ca")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class CRMSyncService : System.Web.Services.WebService
    {
        /// <summary>
        /// TODO
        /// </summary>
        [WebMethod]
        public void UpdateCustomer(System.Xml.XmlDocument customerXML)
        {
            IOrganizationService organizationService = Helper.GetCrmConnection();
            //TODO
        }

        /// <summary>
        /// TODO
        /// </summary>
        [WebMethod]
        public void UpdateContact(System.Xml.XmlDocument contactXML)
        {
            IOrganizationService organizationService = Helper.GetCrmConnection();
            //TODO
        }
    }
}