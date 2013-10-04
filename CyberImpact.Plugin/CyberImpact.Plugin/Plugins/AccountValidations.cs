using System;
using Microsoft.Xrm.Sdk;

namespace CyberImpact.Plugin.Account
{
    public class Validations : IPlugin
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
                        #region Account Validations: Acct Pre Create/Pre Update (1)

                        if ((context.MessageName == Constants.PluginInfo.MessageCreate || context.MessageName == Constants.PluginInfo.MessageUpdate) && context.Stage == (int)Constants.PluginInfo.StageValues.PreValidation)
                        {
                            Entity targetEntity = context.InputParameters["Target"] as Entity;

                            Entity preImage = null;

                            if (context.MessageName == Constants.PluginInfo.MessageUpdate && context.PreEntityImages.Contains("PreImage"))
                            {
                                preImage = context.PreEntityImages["PreImage"] as Entity;
                            }

                            #region if new_Importationcompletee = true, then new_dateImportation must be completed

                            if (targetEntity.Contains(Constants.AccountInfo.Importation_Completee))
                            {
                                if (targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.Importation_Completee))
                                {
                                    if (!targetEntity.Contains(Constants.AccountInfo.Date_Importation))
                                    {
                                        if (preImage == null || (preImage != null && !preImage.Contains(Constants.AccountInfo.Date_Importation)))
                                        {
                                            throw new InvalidPluginExecutionException(OperationStatus.Canceled, "<< " + Constants.AccountInfo.Date_Importation + " is required >>");
                                        }
                                    }
                                }
                            }

                            #endregion if new_Importationcompletee = true, then new_dateImportation must be completed

                            #region if new_gabaritcree = true, then the new_dategabaritcree must be completed

                            if (targetEntity.Contains(Constants.AccountInfo.Gabarit_Crée))
                            {
                                if (targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.Gabarit_Crée))
                                {
                                    if (!targetEntity.Contains(Constants.AccountInfo.Date_Gabarit_Crée))
                                    {
                                        if (preImage == null || (preImage != null && !preImage.Contains(Constants.AccountInfo.Date_Gabarit_Crée)))
                                        {
                                            throw new InvalidPluginExecutionException(OperationStatus.Canceled, "<< " + Constants.AccountInfo.Date_Gabarit_Crée + " is required >>");
                                        }
                                    }
                                }
                            }

                            #endregion if new_gabaritcree = true, then the new_dategabaritcree must be completed

                            #region if new_envoicomplete = true, then the new_dateenvoicomplete

                            if (targetEntity.Contains(Constants.AccountInfo.Envoi_complété))
                            {
                                if (targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.Envoi_complété))
                                {
                                    if (!targetEntity.Contains(Constants.AccountInfo.Date_envoi_complété))
                                    {
                                        if (preImage == null || (preImage != null && !preImage.Contains(Constants.AccountInfo.Date_envoi_complété)))
                                        {
                                            throw new InvalidPluginExecutionException(OperationStatus.Canceled, "<< " + Constants.AccountInfo.Date_envoi_complété + " is required >>");
                                        }
                                    }
                                }
                            }

                            #endregion if new_envoicomplete = true, then the new_dateenvoicomplete

                            #region if new_dateconversioncomptepayant is not null, new_datedechangementdeforfait must be completed

                            if (targetEntity.Contains(Constants.AccountInfo.Date_conversion_compte_payant))
                            {
                                if (!targetEntity.Contains(Constants.AccountInfo.Date_de_changement_de_forfait))
                                {
                                    if (preImage == null || (preImage != null && !preImage.Contains(Constants.AccountInfo.Date_de_changement_de_forfait)))
                                    {
                                        throw new InvalidPluginExecutionException(OperationStatus.Canceled, "<< " + Constants.AccountInfo.Date_de_changement_de_forfait + " is required >>");
                                    }
                                }
                            }

                            #endregion if new_dateconversioncomptepayant is not null, new_datedechangementdeforfait must be completed

                            #region if new_datedechangementdeforfait is not null, new_forfait must be completed

                            if (targetEntity.Contains(Constants.AccountInfo.Date_de_changement_de_forfait))
                            {
                                if (!targetEntity.Contains(Constants.AccountInfo.Forfait))
                                {
                                    if (preImage == null || (preImage != null && !preImage.Contains(Constants.AccountInfo.Forfait)))
                                    {
                                        throw new InvalidPluginExecutionException(OperationStatus.Canceled, "<< " + Constants.AccountInfo.Forfait + " is required >>");
                                    }
                                }
                            }

                            #endregion if new_datedechangementdeforfait is not null, new_forfait must be completed

                            #region If new_supprime becomes true or new_actif becomes false, set the new_datefermeturedecompte to the current date

                            if (targetEntity.Contains(Constants.AccountInfo.BD_supprimée))
                            {
                                if (targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.BD_supprimée))
                                {
                                    targetEntity[Constants.AccountInfo.Date_Fermeture_de_Compte] = DateTime.Today;
                                }
                            }
                            if (targetEntity.Contains(Constants.AccountInfo.Actif))
                            {
                                if (!targetEntity.GetAttributeValue<bool>(Constants.AccountInfo.Actif))
                                {
                                    targetEntity[Constants.AccountInfo.Date_Fermeture_de_Compte] = DateTime.Today;
                                }
                            }

                            #endregion If new_supprime becomes true or new_actif becomes false, set the new_datefermeturedecompte to the current date
                        }

                        #endregion Account Validations: Acct Pre Create/Pre Update (1)
                    }
                }
                catch (Exception ex)
                {
                    _errMsg = "ERROR in Account Validations Plug-in: ";
                    _errMsg += ex.Message;

                    throw new InvalidPluginExecutionException(_errMsg, ex);
                }
            }
        }
    }
}