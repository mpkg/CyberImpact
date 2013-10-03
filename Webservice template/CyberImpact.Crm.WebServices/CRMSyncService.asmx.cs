using System;
using System.ServiceModel;
using System.Web.Services;
using System.Xml;
using CyberImpact.Crm.Operations;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

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
            #region Parse XML

            XmlDocument customerXMLDocument = new XmlDocument();
            try
            {
                customerXMLDocument.LoadXml(customerXML);
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid XML");
            }

            //Create customer object to store XML data.
            Entity customer = new Entity("account");

            XmlNode xmlNode = customerXMLDocument.SelectSingleNode("/Customer");
            if (xmlNode.Attributes.GetNamedItem("NoDuClient") == null)
                throw new Exception("NoDuClient missing from XML");
            else
            {
                try
                {
                    customer["new_idclient"] = Convert.ToInt32(xmlNode.Attributes.GetNamedItem("NoDuClient").Value);

                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/NomDuCompte");
                    if (xmlNode != null)
                        customer["name"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Telephone");
                    if (xmlNode != null)
                        customer["telephone1"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Adresse1");
                    if (xmlNode != null)
                        customer["address1_line1"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Adresse2");
                    if (xmlNode != null)
                        customer["address1_line2"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Ville");
                    if (xmlNode != null)
                        customer["address1_city"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Province");
                    if (xmlNode != null)
                        customer["address1_stateorprovince"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Pays");
                    if (xmlNode != null)
                        customer["address1_country"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/CodePostal");
                    if (xmlNode != null)
                        customer["address1_postalcode"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateDeEssai");
                    if (xmlNode != null)
                        customer["new_datecreation"] = Convert.ToDateTime(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/NombreDeLogin");
                    if (xmlNode != null)
                        customer["new_nombredelogin"] = Convert.ToInt32(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/ImportationCompletee");
                    if (xmlNode != null)
                        customer["new_importationcompletee"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateImportation");
                    if (xmlNode != null)
                        customer["new_dateimportation"] = Convert.ToDateTime(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/GabaritCree");
                    if (xmlNode != null)
                        customer["new_gabaritcree"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateGabarit");
                    if (xmlNode != null)
                        customer["new_dategabaritcree"] = Convert.ToDateTime(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/EnvoiComplete");
                    if (xmlNode != null)
                        customer["new_envoicomplete"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateEnvoi");
                    if (xmlNode != null)
                        customer["new_dateenvoicomplete"] = Convert.ToDateTime(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/DateActivation");
                    if (xmlNode != null)
                        customer["new_dateconversioncomptepayant"] = Convert.ToDateTime(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Forfait");
                    if (xmlNode != null)
                        customer["new_forfait"] = new OptionSetValue(Convert.ToInt32(xmlNode.InnerText));
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/OSBL");
                    if (xmlNode != null)
                        customer["new_osbl"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Premium");
                    if (xmlNode != null)
                        customer["new_comptepremium"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Actif");
                    if (xmlNode != null)
                        customer["new_actif"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/EnvoiPermis");
                    if (xmlNode != null)
                        customer["new_envoipermis"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/Revendeur");
                    if (xmlNode != null)
                        customer["new_revendeur"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/NoDeRevendeur");
                    if (xmlNode != null)
                        customer["new_noderevendeur"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/RevenuForfait");
                    if (xmlNode != null)
                        customer["new_revenuforfait"] = new Money(Convert.ToDecimal(xmlNode.InnerText));
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/RevenuExtra");
                    if (xmlNode != null)
                        customer["new_revenuextra"] = new Money(Convert.ToDecimal(xmlNode.InnerText));
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/CourrielFrom");
                    if (xmlNode != null)
                        customer["new_courrielfrom"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/ExpirationDemo");
                    if (xmlNode != null)
                        customer["new_expirationdudemo"] = Convert.ToDateTime(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/BDSupprimee");
                    if (xmlNode != null)
                        customer["new_supprime"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/AccesAPI");
                    if (xmlNode != null)
                        customer["new_accesapi"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/AccesAPISansOptIn");
                    if (xmlNode != null)
                        customer["new_accesapisansoptin"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/EntrepriseFacturation");
                    if (xmlNode != null)
                        customer["new_entreprisefacturation"] = xmlNode.InnerText;
                    xmlNode = customerXMLDocument.SelectSingleNode("/Customer/MéthodeDePaiement");
                    if (xmlNode != null)
                        customer["new_methodedepaiement"] = new OptionSetValue(Convert.ToInt32(xmlNode.InnerText));
                }
                catch (Exception ex)
                {
                    throw new Exception("Invalid XML");
                }
            }

            #endregion Parse XML

            #region Retrieve customer with NoDuClient value & Update / Create customer accordingly

            try
            {
                IOrganizationService organizationService = Helper.GetCrmConnection();

                QueryExpression query = new QueryExpression("account");
                query.ColumnSet = new ColumnSet(new string[] { "accountid" });
                ConditionExpression condition = new ConditionExpression("new_idclient", ConditionOperator.Equal, customer["new_idclient"]);

                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                filter.AddCondition(condition);

                query.Criteria = filter;

                EntityCollection customers = organizationService.RetrieveMultiple(query);

                if (customers.Entities.Count > 0)//update
                {
                    customer.Id = customers[0].Id;
                    organizationService.Update(customer);
                }
                else //create
                {
                    organizationService.Create(customer);
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("CRM Exception: " + ex);
            }
            catch (Exception ex)
            {
                throw new Exception("CRM Exception: " + ex);
            }

            #endregion Retrieve customer with NoDuClient value & Update / Create customer accordingly
        }

        /// <summary>
        /// Web method to update Contact in CRM
        /// </summary>
        /// <param name="contactXML">Contact information having contact info.</param>
        [WebMethod]
        public void UpdateContact(string contactXML)
        {
            #region Parse XML

            XmlDocument contactXMLDocument = new XmlDocument();
            try
            {
                contactXMLDocument.LoadXml(contactXML);
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid XML");
            }

            //Create customer object to store XML data.
            Entity contact = new Entity("contact");
            int clientId;

            XmlNode xmlNode = contactXMLDocument.SelectSingleNode("/Contact");
            if (xmlNode.Attributes.GetNamedItem("NoDuClient") == null)
                throw new Exception("NoDuClient missing from XML");
            else if (xmlNode.Attributes.GetNamedItem("Nom") == null)
                throw new Exception("Nom missing from XML");
            else if (xmlNode.Attributes.GetNamedItem("Prenom") == null)
                throw new Exception("Prenom missing from XML");
            else
            {
                try
                {
                    clientId = Convert.ToInt32(xmlNode.Attributes.GetNamedItem("NoDuClient").Value);
                    contact["lastname"] = xmlNode.Attributes.GetNamedItem("Nom").Value;
                    contact["firstname"] = xmlNode.Attributes.GetNamedItem("Prenom").Value;

                    xmlNode = contactXMLDocument.SelectSingleNode("/Contact/ContactTechnique");
                    if (xmlNode != null)
                        contact["new_contacttechnique"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = contactXMLDocument.SelectSingleNode("/Contact/ContactFacturation");
                    if (xmlNode != null)
                        contact["new_contactfacturation"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = contactXMLDocument.SelectSingleNode("/Contact/CyberBulletin");
                    if (xmlNode != null)
                        contact["new_cyberbulletin"] = Convert.ToBoolean(xmlNode.InnerText);
                    xmlNode = contactXMLDocument.SelectSingleNode("/Contact/Langue");
                    if (xmlNode != null)
                        contact["new_languecyberbulletin"] = new OptionSetValue(Convert.ToInt32(xmlNode.InnerText));
                    xmlNode = contactXMLDocument.SelectSingleNode("/Contact/Actif");
                    if (xmlNode != null)
                        contact["new_actif"] = Convert.ToBoolean(xmlNode.InnerText);
                }
                catch (Exception ex)
                {
                    throw new Exception("Invalid XML");
                }
            }

            #endregion Parse XML

            #region Retrieve contact with NoDuClient, Prenom and Nom & Update / Create accordingly.

            try
            {
                IOrganizationService organizationService = Helper.GetCrmConnection();

                QueryExpression query1 = new QueryExpression("account");
                query1.ColumnSet = new ColumnSet(new string[] { "accountid" });
                ConditionExpression condition1 = new ConditionExpression("new_idclient", ConditionOperator.Equal, clientId);

                FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                filter1.AddCondition(condition1);

                query1.Criteria = filter1;

                EntityCollection customers = organizationService.RetrieveMultiple(query1);

                if (customers.Entities.Count > 0)
                {
                    QueryExpression query2 = new QueryExpression("contact");
                    query2.ColumnSet = new ColumnSet(new string[] { "contactid" });

                    ConditionExpression condition21 = new ConditionExpression("parentcustomerid", ConditionOperator.Equal, customers[0].Id);
                    ConditionExpression condition22 = new ConditionExpression("lastname", ConditionOperator.Equal, contact["lastname"]);
                    ConditionExpression condition23 = new ConditionExpression("firstname", ConditionOperator.Equal, contact["firstname"]);

                    FilterExpression filter2 = new FilterExpression(LogicalOperator.And);
                    filter2.AddCondition(condition21);
                    filter2.AddCondition(condition22);
                    filter2.AddCondition(condition23);

                    query2.Criteria = filter2;

                    EntityCollection contacts = organizationService.RetrieveMultiple(query2);
                    if (contacts.Entities.Count > 0)//update contact
                    {
                        contact.Id = contacts[0].Id;
                        organizationService.Update(contact);
                    }
                    else //create contact
                    {
                        contact["parentcustomerid"] = new EntityReference("account", customers[0].Id);
                        organizationService.Create(contact);
                    }
                }
                else
                {
                    throw new Exception("No account found having specified NoDuClient");
                }

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception("CRM Exception: " + ex);
            }
            catch (Exception ex)
            {
                throw new Exception("CRM Exception: " + ex);
            }

            #endregion Retrieve contact with NoDuClient, Prenom and Nom & Update / Create accordingly.
        }
    }
}