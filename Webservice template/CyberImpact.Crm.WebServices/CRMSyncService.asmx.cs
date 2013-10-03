using System;
using System.Web.Services;
using System.Xml;
using CyberImpact.Crm.Operations;
using Microsoft.Xrm.Sdk;

namespace CyberImpact.Crm.WebServices
{
    /// <summary>
    /// Web Services to update customer & contact data in CRM
    /// </summary>
    [WebService(Namespace = "http://cyberimpact.ca")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class CRMSyncService : System.Web.Services.WebService
    {
        /// <summary>
        /// Web method to update Customer in CRM
        /// </summary>
        /// <param name="customerXML">XML document having customer info.</param>
        [WebMethod]
        public void UpdateCustomer(string customerXML)
        {
            IOrganizationService organizationService = Helper.GetCrmConnection();

            #region Parse XML

            try
            {
                //Create customer object to store XML data.
                Customer customer = new Customer();

                XmlDocument customerXMLDocument = new XmlDocument();
                customerXMLDocument.LoadXml(customerXML);

                XmlNode xmlNode = customerXMLDocument.SelectSingleNode("/Customer");
                customer.noDuClient = Convert.ToInt32(xmlNode.Attributes.GetNamedItem("NoDuClient").Value);

                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/NomDuCompte");
                if (xmlNode != null)
                    customer.nomDuCompte = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Telephone");
                if (xmlNode != null)
                    customer.telephone = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Adresse1");
                if (xmlNode != null)
                    customer.adresse1 = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Adresse2");
                if (xmlNode != null)
                    customer.adresse2 = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Ville");
                if (xmlNode != null)
                    customer.ville = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Province");
                if (xmlNode != null)
                    customer.province = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Pays");
                if (xmlNode != null)
                    customer.pays = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/CodePostal");
                if (xmlNode != null)
                    customer.codePostal = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateDeEssai");
                if (xmlNode != null)
                    customer.dateDeEssai = Convert.ToDateTime(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/NombreDeLogin");
                if (xmlNode != null)
                    customer.nombreDeLogin = Convert.ToInt32(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/ImportationCompletee");
                if (xmlNode != null)
                    customer.importationCompletee = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateImportation");
                if (xmlNode != null)
                    customer.dateImportation = Convert.ToDateTime(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/GabaritCree");
                if (xmlNode != null)
                    customer.gabaritCree = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateGabarit");
                if (xmlNode != null)
                    customer.dateGabarit = Convert.ToDateTime(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/EnvoiComplete");
                if (xmlNode != null)
                    customer.envoiComplete = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateEnvoi");
                if (xmlNode != null)
                    customer.dateEnvoi = Convert.ToDateTime(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateActivation");
                if (xmlNode != null)
                    customer.dateActivation = Convert.ToDateTime(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Forfait");
                if (xmlNode != null)
                    customer.forfait = Convert.ToInt32(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/OSBL");
                if (xmlNode != null)
                    customer.osbl = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Premium");
                if (xmlNode != null)
                    customer.premium = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Actif");
                if (xmlNode != null)
                    customer.actif = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/EnvoiPermis");
                if (xmlNode != null)
                    customer.envoiPermis = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Revendeur");
                if (xmlNode != null)
                    customer.revendeur = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/NoDeRevendeur");
                if (xmlNode != null)
                    customer.noDeRevendeur = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/RevenuForfait");
                if (xmlNode != null)
                    customer.revenuForfait = new Money(Convert.ToDecimal(xmlNode.InnerText));
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/RevenuExtra");
                if (xmlNode != null)
                    customer.revenuExtra = new Money(Convert.ToDecimal(xmlNode.InnerText));
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/CourrielFrom");
                if (xmlNode != null)
                    customer.courrielFrom = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/ExpirationDemo");
                if (xmlNode != null)
                    customer.expirationDemo = Convert.ToDateTime(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/BDSupprimee");
                if (xmlNode != null)
                    customer.bDSupprimee = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/AccesAPI");
                if (xmlNode != null)
                    customer.accesAPI = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/AccesAPISansOptIn");
                if (xmlNode != null)
                    customer.accesAPISansOptIn = Convert.ToBoolean(xmlNode.InnerText);
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/EntrepriseFacturation");
                if (xmlNode != null)
                    customer.entrepriseFacturation = xmlNode.InnerText;
                xmlNode = customerXMLDocument.SelectSingleNode("/Customer/MéthodeDePaiement");
                if (xmlNode != null)
                    customer.methodeDePaiement = Convert.ToInt32(xmlNode.InnerText);
            }
            catch (XmlException ex)
            {
                throw new Exception("Invalid XML" + ex.Message);
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid XML" + ex.Message);
            }

            #endregion Parse XML

            #region Retrieve customer with NoDuClient value

            #endregion Retrieve customer with NoDuClient value

            #region Update / Create customer

            #endregion Update / Create customer
        }

        /// <summary>
        /// Web method to update Contact in CRM
        /// </summary>
        /// <param name="contactXML">Contact information having contact info.</param>
        [WebMethod]
        public void UpdateContact(string contactXML)
        {
            IOrganizationService organizationService = Helper.GetCrmConnection();
            //TODO
        }
    }
}