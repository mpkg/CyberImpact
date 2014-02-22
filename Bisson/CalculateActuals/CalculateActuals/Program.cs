using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CalculateActuals
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            IOrganizationService organizationService = Helper.GetCrmConnection();

            #region Retrieve active Budget records

            QueryExpression budgetQuery = new QueryExpression()
            {
                EntityName = "new_budget",
                ColumnSet = new ColumnSet(new string[] { "new_objectif", "new_vendeur", "new_typedetravaux", "new_periodefiscale", "new_excercice" }),
                Criteria = new FilterExpression()
                {
                    Conditions =
                    {
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                    }
                }
            };
            EntityCollection budgetRecords = organizationService.RetrieveMultiple(budgetQuery);

            #endregion Retrieve active Budget records

            decimal objectif;
            EntityReference vendeur;
            EntityReference typedetravaux;
            int fiscalperiod;
            int exercice;

            for (int i = 0; i < budgetRecords.Entities.Count; i++)
            {
                objectif = budgetRecords.Entities[i].GetAttributeValue<decimal>("new_objectif");
                vendeur = budgetRecords.Entities[i].GetAttributeValue<EntityReference>("new_vendeur");
                typedetravaux = budgetRecords.Entities[i].GetAttributeValue<EntityReference>("new_typedetravaux");
                fiscalperiod = budgetRecords.Entities[i].GetAttributeValue<OptionSetValue>("new_periodefiscale").Value;
                exercice = budgetRecords.Entities[i].GetAttributeValue<OptionSetValue>("new_excercice").Value;

                QueryExpression contractQuery1 = new QueryExpression()
                {
                    EntityName = "new_contracttravaux",
                    ColumnSet = new ColumnSet("new_ventetotal"),
                    Criteria = new FilterExpression()
                   {
                       Conditions =
                    {
                        new ConditionExpression("new_vendeur", ConditionOperator.Equal, vendeur),
                        new ConditionExpression("new_typedetravaux", ConditionOperator.Equal, typedetravaux),
                        new ConditionExpression("new_datecontrat", ConditionOperator.InFiscalPeriod, fiscalperiod),
                        new ConditionExpression("new_datecontrat", ConditionOperator.InFiscalYear, exercice)
                    }
                   }
                };
                EntityCollection contractRecords = organizationService.RetrieveMultiple(contractQuery1);
                for (int j = 0; j < contractRecords.Entities.Count; j++)
                { 
                    Console.WriteLine(); 
                }
                
            }

            Console.Read();
        }
    }
}