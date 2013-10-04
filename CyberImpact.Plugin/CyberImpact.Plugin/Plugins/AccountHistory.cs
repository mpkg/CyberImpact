using System;
using Microsoft.Xrm.Sdk;

namespace CyberImpact.Plugin.Account
{
    public class History : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.PrimaryEntityName == Constants.AccountInfo.EntityName)
            {
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService callingUserService = factory.CreateOrganizationService(context.InitiatingUserId);

                string _errMsg = "";

                try
                {
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        #region Account History: Acct Post Create/Pre Update (1)

                        if ((context.MessageName == Constants.PluginInfo.MessageCreate || context.MessageName == Constants.PluginInfo.MessageUpdate) && context.Stage == (int)Constants.PluginInfo.StageValues.PostOperation)
                        {
                            Entity targetEntity = context.InputParameters["Target"] as Entity;

                            Entity postImage = null;

                            if (context.MessageName == Constants.PluginInfo.MessageUpdate && context.PostEntityImages.Contains("PostImage"))
                            {
                                postImage = context.PostEntityImages["PostImage"] as Entity;
                            }

                            #region If new_forfait, new_comptepremium or new_osbl changes, create a new record in the new_forfaithistorique

                            if (targetEntity.Contains(Constants.AccountInfo.Forfait) || targetEntity.Contains(Constants.AccountInfo.Compte_premium) || targetEntity.Contains(Constants.AccountInfo.OSBL))
                            {
                                Entity forfaithistorique = new Entity(Constants.ForfaitHistoriqueInfo.EntityName);
                                forfaithistorique[Constants.ForfaitHistoriqueInfo.Date_d_activation] = DateTime.Now;
                                forfaithistorique[Constants.ForfaitHistoriqueInfo.Compte] = new EntityReference(Constants.AccountInfo.EntityName, targetEntity.Id);

                                if (context.MessageName == Constants.PluginInfo.MessageCreate)
                                {
                                    forfaithistorique[Constants.ForfaitHistoriqueInfo.Forfait] = targetEntity.GetAttributeValue<OptionSetValue>(Constants.AccountInfo.Forfait);
                                    forfaithistorique[Constants.ForfaitHistoriqueInfo.OSBL] = targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.OSBL);
                                    forfaithistorique[Constants.ForfaitHistoriqueInfo.Premium] = targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.Compte_premium);
                                }
                                else if (context.MessageName == Constants.PluginInfo.MessageUpdate && postImage != null)
                                {
                                    forfaithistorique[Constants.ForfaitHistoriqueInfo.Forfait] = postImage.GetAttributeValue<OptionSetValue>(Constants.AccountInfo.Forfait);
                                    forfaithistorique[Constants.ForfaitHistoriqueInfo.OSBL] = postImage.GetAttributeValue<bool>(Constants.AccountInfo.OSBL);
                                    forfaithistorique[Constants.ForfaitHistoriqueInfo.Premium] = postImage.GetAttributeValue<bool>(Constants.AccountInfo.Compte_premium);
                                }
                                callingUserService.Create(forfaithistorique);
                            }

                            #endregion If new_forfait, new_comptepremium or new_osbl changes, create a new record in the new_forfaithistorique

                            #region If new_revenuforfait or new_revenuextra changes, create a new record in the new_revenuhistorique

                            if (targetEntity.Contains(Constants.AccountInfo.Revenu_Forfait) || targetEntity.Contains(Constants.AccountInfo.Revenu_Extra))
                            {
                                Entity revenuhistorique = new Entity(Constants.RevenuHistoriqueInfo.EntityName);
                                revenuhistorique[Constants.RevenuHistoriqueInfo.Date_du_Revenu] = DateTime.Now;
                                revenuhistorique[Constants.RevenuHistoriqueInfo.Compte] = new EntityReference(Constants.AccountInfo.EntityName, targetEntity.Id);

                                if (context.MessageName == Constants.PluginInfo.MessageCreate)
                                {
                                    revenuhistorique[Constants.RevenuHistoriqueInfo.Revenu_Extra] = targetEntity.GetAttributeValue<Money>(Constants.AccountInfo.Revenu_Extra);
                                    revenuhistorique[Constants.RevenuHistoriqueInfo.Revenu_Forfait] = targetEntity.GetAttributeValue<Money>(Constants.AccountInfo.Revenu_Forfait);
                                }
                                else if (context.MessageName == Constants.PluginInfo.MessageUpdate)
                                {
                                    revenuhistorique[Constants.RevenuHistoriqueInfo.Revenu_Extra] = postImage.GetAttributeValue<Money>(Constants.AccountInfo.Revenu_Extra);
                                    revenuhistorique[Constants.RevenuHistoriqueInfo.Revenu_Forfait] = postImage.GetAttributeValue<Money>(Constants.AccountInfo.Revenu_Forfait);
                                }
                                callingUserService.Create(revenuhistorique);
                            }

                            #endregion If new_revenuforfait or new_revenuextra changes, create a new record in the new_revenuhistorique
                        }

                        #endregion Account History: Acct Pre Create/Pre Update (1)
                    }
                }
                catch (Exception ex)
                {
                    _errMsg = "ERROR in Account History Plug-in: ";
                    _errMsg += ex.Message;

                    throw new InvalidPluginExecutionException(_errMsg, ex);
                }
            }
        }
    }
}