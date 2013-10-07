using System;
using Microsoft.Xrm.Sdk;

namespace CyberImpact.Plugin.Entity
{
    public class Setname : IPlugin
    {
        private IOrganizationService callingUserService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            callingUserService = factory.CreateOrganizationService(context.InitiatingUserId);

            string _errMsg = "";

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Microsoft.Xrm.Sdk.Entity)
                {
                    #region Entity Set name: new_forfaithistorique, new_revenuhistorique Pre Create/Update

                    if ((context.MessageName == Constants.PluginInfo.MessageCreate || context.MessageName == Constants.PluginInfo.MessageUpdate) && context.Stage == (int)Constants.PluginInfo.StageValues.PreValidation)
                    {
                        Microsoft.Xrm.Sdk.Entity targetEntity = context.InputParameters["Target"] as Microsoft.Xrm.Sdk.Entity;

                        Microsoft.Xrm.Sdk.Entity preImage = null;

                        if (context.MessageName == Constants.PluginInfo.MessageUpdate && context.PreEntityImages.Contains("PreImage"))
                        {
                            preImage = context.PreEntityImages["PreImage"] as Microsoft.Xrm.Sdk.Entity;
                        }

                        EntityReference compte = null;
                        DateTime? tempDate;
                        string formattedDateValue = string.Empty;

                        #region Forfait Historique

                        if (context.PrimaryEntityName == Constants.ForfaitHistoriqueInfo.EntityName)
                        {
                            //check if target contains compte, else fetch compte from image
                            if (targetEntity.Contains(Constants.ForfaitHistoriqueInfo.Compte))
                            {
                                compte = targetEntity.GetAttributeValue<EntityReference>(Constants.ForfaitHistoriqueInfo.Compte);
                            }

                            else if (preImage.Contains(Constants.ForfaitHistoriqueInfo.Compte))
                            {
                                compte = preImage.GetAttributeValue<EntityReference>(Constants.ForfaitHistoriqueInfo.Compte);
                            }

                            //check if target contains date, else fetch date from image
                            if (targetEntity.Contains(Constants.ForfaitHistoriqueInfo.Date_d_activation))
                            {
                                tempDate = targetEntity.GetAttributeValue<DateTime?>(Constants.ForfaitHistoriqueInfo.Date_d_activation);
                                if (tempDate.HasValue)
                                    formattedDateValue = tempDate.Value.ToString("ddMMyyyy");
                            }
                            else if (preImage.Contains(Constants.ForfaitHistoriqueInfo.Date_d_activation))
                            {
                                tempDate = preImage.GetAttributeValue<DateTime?>(Constants.ForfaitHistoriqueInfo.Date_d_activation);
                                if (tempDate.HasValue)
                                    formattedDateValue = tempDate.Value.ToString("ddMMyyyy");
                            }

                            //set name field
                            targetEntity[Constants.ForfaitHistoriqueInfo.Name] = (compte != null ? GetAccountName(compte.Id) : "") + " " + formattedDateValue;
                        }

                        #endregion Forfait Historique

                        #region Revenu Historique

                        if (context.PrimaryEntityName == Constants.RevenuHistoriqueInfo.EntityName)
                        {
                            //check if target contains compte, else fetch compte from image
                            if (targetEntity.Contains(Constants.RevenuHistoriqueInfo.Compte))
                            {
                                compte = targetEntity.GetAttributeValue<EntityReference>(Constants.RevenuHistoriqueInfo.Compte);
                            }

                            else if (preImage.Contains(Constants.RevenuHistoriqueInfo.Compte))
                            {
                                compte = preImage.GetAttributeValue<EntityReference>(Constants.RevenuHistoriqueInfo.Compte);
                            }

                            //check if target contains date, else fetch date from image
                            if (targetEntity.Contains(Constants.RevenuHistoriqueInfo.Date_du_Revenu))
                            {
                                tempDate = targetEntity.GetAttributeValue<DateTime?>(Constants.RevenuHistoriqueInfo.Date_du_Revenu);
                                if (tempDate.HasValue)
                                    formattedDateValue = tempDate.Value.ToString("ddMMyyyy");
                            }
                            else if (preImage.Contains(Constants.RevenuHistoriqueInfo.Date_du_Revenu))
                            {
                                tempDate = preImage.GetAttributeValue<DateTime?>(Constants.RevenuHistoriqueInfo.Date_du_Revenu);
                                if (tempDate.HasValue)
                                    formattedDateValue = tempDate.Value.ToString("ddMMyyyy");
                            }

                            //set name field
                            targetEntity[Constants.RevenuHistoriqueInfo.Name] = (compte != null ? GetAccountName(compte.Id) : "") + " " + formattedDateValue;
                        }

                        #endregion Revenu Historique
                    }

                    #endregion Entity Set name: new_forfaithistorique, new_revenuhistorique Pre Create/Update
                }
            }
            catch (Exception ex)
            {
                _errMsg = "ERROR in Entity Setname Plug-in: ";
                _errMsg += ex.Message;

                throw new InvalidPluginExecutionException(_errMsg, ex);
            }
        }

        public string GetAccountName(Guid accountId)
        {
            Microsoft.Xrm.Sdk.Entity account = callingUserService.Retrieve(Constants.AccountInfo.EntityName, accountId, new Microsoft.Xrm.Sdk.Query.ColumnSet(new string[] { Constants.AccountInfo.AccountName }));
            return (string)account[Constants.AccountInfo.AccountName];
        }
    }
}