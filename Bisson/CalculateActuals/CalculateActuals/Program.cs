using System;
using Microsoft.Xrm.Sdk;

using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace CalculateActuals
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            #region Variables

            decimal objectif = decimal.Zero;
            EntityReference vendeur;
            EntityReference typedetravaux;
            int fiscalperiod;
            int exercice;

            EntityCollection budgetRecords;
            List<Entity> budgetRecordsToUpdate = new List<Entity>();
            EntityCollection contractRecordsAC;
            EntityCollection contractRecordsAP;

            decimal reelAC;
            decimal reelAP;
            decimal pctAC = decimal.Zero;

            #endregion            
            
            try
            {
                Console.WriteLine("Calculating Budget Actuals...");

                IOrganizationService organizationService = Helper.GetCrmConnection();

                Console.WriteLine("\nOrganisation service object created.");

                #region Retrieve Active Budget Records

                QueryExpression budgetQuery = new QueryExpression()
                {
                    EntityName = "new_budget",
                    ColumnSet = new ColumnSet(new string[] { "new_objectif", "new_vendeur", "new_typedetravaux", "new_periodefiscale", "new_excercice" }),
                    Criteria = new FilterExpression(LogicalOperator.And)
                    {
                        Conditions =
                    {
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                        new ConditionExpression("createdon", ConditionOperator.Today)
                    }
                    }
                };

                budgetRecords = organizationService.RetrieveMultiple(budgetQuery);

                #endregion Retrieve Active Budget Records

                Console.WriteLine("\nRetrieved " + budgetRecords.Entities.Count + " active budget records.");

                #region Process each budget record

                for (int i = 0; i < budgetRecords.Entities.Count; i++)
                {
                    try
                    {
                        objectif = budgetRecords.Entities[i].GetAttributeValue<Money>("new_objectif").Value;
                        vendeur = budgetRecords.Entities[i].GetAttributeValue<EntityReference>("new_vendeur");
                        typedetravaux = budgetRecords.Entities[i].GetAttributeValue<EntityReference>("new_typedetravaux");
                        fiscalperiod = budgetRecords.Entities[i].GetAttributeValue<OptionSetValue>("new_periodefiscale").Value;
                        exercice = budgetRecords.Entities[i].GetAttributeValue<OptionSetValue>("new_excercice").Value;

                        Console.WriteLine("\nBudget record - " + vendeur.Name + "/" + typedetravaux.Name + "/" + fiscalperiod.ToString() + "/" + exercice.ToString());

                        #region Reel AC Calculation

                        QueryExpression contractQueryAC = new QueryExpression()
                        {
                            EntityName = "new_contrat",
                            ColumnSet = new ColumnSet("new_ventetotal"),
                            Criteria = new FilterExpression(LogicalOperator.And)
                            {
                                Conditions =
                                {
                                    new ConditionExpression("new_vendeur1", ConditionOperator.Equal, vendeur.Id),
                                    new ConditionExpression("new_typedetravaux1", ConditionOperator.Equal, typedetravaux.Id),
                                    new ConditionExpression("new_datecontrat", ConditionOperator.InFiscalPeriod, fiscalperiod-100),
                                    new ConditionExpression("new_datecontrat", ConditionOperator.InFiscalYear, exercice)
                                }
                            }
                        };
                        contractRecordsAC = organizationService.RetrieveMultiple(contractQueryAC);

                        Console.WriteLine("Retrieved " + contractRecordsAC.Entities.Count + " contract records for Reel AC calculation.");

                        reelAC = decimal.Zero;
                        for (int j = 0; j < contractRecordsAC.Entities.Count; j++)
                        {
                            reelAC += contractRecordsAC.Entities[j].GetAttributeValue<Money>("new_ventetotal").Value;
                        }
                        budgetRecords.Entities[i]["new_reelac"] = new Money(reelAC);
                        #endregion Reel AC Calculation

                        #region Reel AP Calculation

                        QueryExpression contractQueryAP = new QueryExpression()
                        {
                            EntityName = "new_contrat",
                            ColumnSet = new ColumnSet("new_ventetotal"),
                            Criteria = new FilterExpression(LogicalOperator.And)
                            {
                                Conditions =
                                {
                                     new ConditionExpression("new_vendeur1", ConditionOperator.Equal, vendeur.Id),
                                     new ConditionExpression("new_typedetravaux1", ConditionOperator.Equal, typedetravaux.Id),
                                     new ConditionExpression("new_datecontrat", ConditionOperator.InFiscalPeriod, fiscalperiod-100),
                                     new ConditionExpression("new_datecontrat", ConditionOperator.InFiscalYear, exercice-1)
                                }
                            }
                        };
                        contractRecordsAP = organizationService.RetrieveMultiple(contractQueryAP);

                        Console.WriteLine("Retrieved " + contractRecordsAP.Entities.Count + " contract records for Reel AP calculation.");

                        reelAP = decimal.Zero;
                        for (int j = 0; j < contractRecordsAP.Entities.Count; j++)
                        {
                            reelAP += contractRecordsAP.Entities[j].GetAttributeValue<Money>("new_ventetotal").Value;
                        }

                        budgetRecords.Entities[i]["new_reelap"] = new Money(reelAP);
                        #endregion Reel AP Calculation

                        #region PCT AC Calculation
                        if (objectif != 0)
                        {
                            pctAC = reelAC / objectif * 100;
                            budgetRecords.Entities[i]["new_pctac"] = pctAC;
                        }

                        #endregion PCT AC Calculation

                        #region Remove other attributes
                        if (budgetRecords.Entities[i].Contains("new_objectif"))
                            budgetRecords.Entities[i].Attributes.Remove("new_objectif");
                        if (budgetRecords.Entities[i].Contains("new_vendeur"))
                            budgetRecords.Entities[i].Attributes.Remove("new_vendeur");
                        if (budgetRecords.Entities[i].Contains("new_typedetravaux"))
                            budgetRecords.Entities[i].Attributes.Remove("new_typedetravaux");
                        if (budgetRecords.Entities[i].Contains("new_periodefiscale"))
                            budgetRecords.Entities[i].Attributes.Remove("new_periodefiscale");
                        if (budgetRecords.Entities[i].Contains("new_excercice"))
                            budgetRecords.Entities[i].Attributes.Remove("new_excercice");

                        #endregion

                        budgetRecordsToUpdate.Add(budgetRecords.Entities[i]);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("For loop - " + ex.Message);
                    }
                }

                #endregion
                
                Console.WriteLine("\nUpdating budget records...");

                Helper.BulkUpdate(organizationService, budgetRecordsToUpdate);

                Console.WriteLine("\nProcess complete!");
            }
            catch (Exception e)
            {
                Console.WriteLine("Main - " + e.Message);
            }
            Console.Read();
        }
    }
}