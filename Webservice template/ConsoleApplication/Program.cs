namespace ConsoleApplication
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string contactXML = @"<Contact NoDuClient='50' Nom='ab v vvc' Prenom='pqr'>
                                    <ContactTechnique>true</ContactTechnique>
                                    <ContactFacturation>false</ContactFacturation>
                                    <Langue>100000000</Langue>
                                    <Actif>true</Actif>
                                </Contact>";

            string customerXML = @"<Customer NoDuClient='50'>
                                    <NomDuCompte>abcFieldValue</NomDuCompte>
                                    <Telephone>FieldValue</Telephone>
                                    <Adresse1>FieldValue</Adresse1>
                                    <Adresse2>FieldValue</Adresse2>
                                    <Ville>FieldValue</Ville>
                                    <Province>FieldValue</Province>
                                    <Pays>FieldValue</Pays>
                                    <CodePostal>FieldValue</CodePostal>
                                    <DateDeEssai>10/10/2013</DateDeEssai>
                                    <NombreDeLogin>1</NombreDeLogin>
                                    <ImportationCompletee>true</ImportationCompletee>
                                    <DateImportation>10/10/2013</DateImportation>
                                    <GabaritCree>true</GabaritCree>
                                    <DateGabarit>10/10/2013</DateGabarit>
                                    <EnvoiComplete>true</EnvoiComplete>
                                    <DateEnvoi>10/10/2013</DateEnvoi>
                                    <DateActivation>10/10/2013</DateActivation>
                                    <Forfait>100000002</Forfait>
                                    <OSBL>true</OSBL>
                                    <Premium>true</Premium>
                                    <Actif>true</Actif>
                                    <EnvoiPermis>true</EnvoiPermis>
                                    <Revendeur>true</Revendeur>
                                    <NoDeRevendeur>FieldValue</NoDeRevendeur>
                                    <RevenuForfait>1445.99</RevenuForfait>
                                    <RevenuExtra>1567.76</RevenuExtra>
                                    <CourrielFrom>FieldValue</CourrielFrom>
                                    <ExpirationDemo>10/10/2013</ExpirationDemo>
                                    <BDSupprimee>true</BDSupprimee>
                                    <AccesAPI>true</AccesAPI>
                                    <AccesAPISansOptIn>true</AccesAPISansOptIn>
                                    <EntrepriseFacturation>FieldValue</EntrepriseFacturation>
                                    <MéthodeDePaiement>100000002</MéthodeDePaiement>
                                </Customer>";

            CRMService.CRMSyncServiceSoapClient client = new CRMService.CRMSyncServiceSoapClient();
            //client.UpdateCustomer(customerXML);
            client.UpdateContact(contactXML);
        }
    }
}