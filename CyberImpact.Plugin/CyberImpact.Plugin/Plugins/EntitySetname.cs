using System;
using Microsoft.Xrm.Sdk;

namespace CyberImpact.Plugin.CustomEntity
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    #region Entity Set name: new_forfaithistorique, new_revenuhistorique Pre Create

                    if (context.MessageName == Constants.PluginInfo.MessageCreate && context.Stage == (int)Constants.PluginInfo.StageValues.PreValidation)
                    {
                        Entity targetEntity = context.InputParameters["Target"] as Entity;
                        if (context.PrimaryEntityName == Constants.ForfaitHistoriqueInfo.EntityName)
                        {
                            targetEntity[Constants.ForfaitHistoriqueInfo.Name] =
                                GetAccountName(targetEntity.GetAttributeValue<EntityReference>(Constants.ForfaitHistoriqueInfo.Compte).Id) + " " +
                                targetEntity.GetAttributeValue<DateTime>(Constants.ForfaitHistoriqueInfo.Date_d_activation).ToString("ddMMyyyy");
                        }
                        if (context.PrimaryEntityName == Constants.RevenuHistoriqueInfo.EntityName)
                        {
                            targetEntity[Constants.RevenuHistoriqueInfo.Name] =
                                GetAccountName(targetEntity.GetAttributeValue<EntityReference>(Constants.RevenuHistoriqueInfo.Compte).Id) + " " +
                                targetEntity.GetAttributeValue<DateTime>(Constants.RevenuHistoriqueInfo.Date_du_Revenu).ToString("ddMMyyyy");
                        }
                    }

                    #endregion Entity Set name: new_forfaithistorique, new_revenuhistorique Pre Create
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
            Entity account = callingUserService.Retrieve(Constants.AccountInfo.EntityName, accountId, new Microsoft.Xrm.Sdk.Query.ColumnSet(new string[] { Constants.AccountInfo.AccountName }));
            return (string)account[Constants.AccountInfo.AccountName];
        }
    }
}