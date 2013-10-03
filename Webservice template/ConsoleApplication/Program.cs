using System;
using System.Xml;
using CyberImpact.Crm.Operations;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsoleApplication
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string customerXML = @"<Customer NoDuClient='10'>
                                        <NomDuCompte>FieldValue</NomDuCompte>
                                        <Telephone>FieldValue</Telephone>
                                        <Adresse1>FieldValue</Adresse1>
                                        <Adresse2>FieldValue</Adresse2>
                                        <Ville>FieldValue</Ville>
                                        <Province>FieldValue</Province>
                                        <Pays>FieldValue</Pays>
                                        <CodePostal>FieldValue</CodePostal>
                                        <DateDeEssai>10/3/2013</DateDeEssai>
                                        <NombreDeLogin>10</NombreDeLogin>
                                        <ImportationCompletee>true</ImportationCompletee>
                                        <DateImportation>10/3/2013</DateImportation>
                                        <GabaritCree>true</GabaritCree>
                                        <DateGabarit>10/3/2013</DateGabarit>
                                        <EnvoiComplete>true</EnvoiComplete>
                                        <DateEnvoi>10/3/2013</DateEnvoi>
                                        <DateActivation>10/3/2013</DateActivation>
                                        <Forfait>2</Forfait>
                                        <OSBL>true</OSBL>
                                        <Premium>true</Premium>
                                        <Actif>true</Actif>
                                        <EnvoiPermis>true</EnvoiPermis>
                                        <Revendeur>true</Revendeur>
                                        <NoDeRevendeur>FieldValue</NoDeRevendeur>
                                        <RevenuForfait>100</RevenuForfait>
                                        <RevenuExtra>1000</RevenuExtra>
                                        <CourrielFrom>10/3/2013</CourrielFrom>
                                        <ExpirationDemo>10/3/2013</ExpirationDemo>
                                        <BDSupprimee>true</BDSupprimee>
                                        <AccesAPI>true</AccesAPI>
                                        <AccesAPISansOptIn>true</AccesAPISansOptIn>
                                        <EntrepriseFacturation>FieldValue</EntrepriseFacturation>
                                        <MéthodeDePaiement>1</MéthodeDePaiement>
                                    </Customer>";

            #region Parse XML

            //Create customer object to store XML data.
            Entity customer = new Entity("account");

            XmlDocument customerXMLDocument = new XmlDocument();
            try
            {
                customerXMLDocument.LoadXml(customerXML);
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid XML");
            }

            XmlNode xmlNode = customerXMLDocument.SelectSingleNode("/Customer");
            if (xmlNode.Attributes.GetNamedItem("NoDuClient") == null)
                throw new Exception("NoDuClient missing from XML");
            else
            {
                try
                {
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
                catch (Exception ex)
                {
                    throw new Exception("Invalid XML");
                }
            }

            #endregion Parse XML

            #region Retrieve customer with NoDuClient value

            IOrganizationService organizationService = Helper.GetCrmConnection();

            QueryExpression query = new QueryExpression("account");
            query.ColumnSet = new ColumnSet(new string[] { "accountid" });
            ConditionExpression condition = new ConditionExpression("new_idclient", ConditionOperator.Equal, customer.noDuClient);

            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.AddCondition(condition);

            query.Criteria = filter;

            EntityCollection accounts = organizationService.RetrieveMultiple(query);
            Entity account = null;
            if (accounts.Entities.Count > 0)
                account = accounts[0];
                

            #endregion Retrieve customer with NoDuClient value

            #region Update / Create customer
            if (account != null)//update 
            {

            }
            else if (account == null)//create
            {

            }
            #endregion Update / Create customer
        }
    }
}