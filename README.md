using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CXPPlugins
{       //Added a step: On Update of CR Industry field from Account update the CR Industry field in all related opportunities 
    public class PushNewClientNumberToRelatedOppsAndMatters : IPlugin
    {
        //private IOrganizationService objService;
        //private ITracingService objtracer;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            ITracingService objtracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.UserId);

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                // Verify that the target entity represents an entity type you are expecting. 
                // For example, an account. If not, the plug-in was not registered correctly.
                //if (entity.LogicalName != "opportunity" && entity.LogicalName != "connection")
                if (entity.LogicalName != "account")
                {
                    //tracer.Trace("entity is not opportunity or connection");
                    objtracer.Trace("entity is not account");
                    return;
                }

                PluginUtilities pluginUtilities = new PluginUtilities();
                //pluginUtilities.Service = objService;
                //pluginUtilities.Tracer = objtracer;

                try
                {
                    Entity account = (Entity)context.InputParameters["Target"];

                    if (account.Contains("accountnumber"))
                    {
                        string accountNumber = account.GetAttributeValue<string>("accountnumber");
                        Entity dbAccount = pluginUtilities.GetAccount(account.Id, objService, objtracer);
                        List<Entity> opportunities = pluginUtilities.GetAllOpportunitiesForAccount(account.Id, objService, objtracer);

                        foreach (Entity currOpportunity in opportunities)
                        {
                            Entity modifiedOpportunity = new Entity("opportunity", currOpportunity.Id);
                            modifiedOpportunity["cxp_clientnumber"] = accountNumber;
                            if (dbAccount != null &&
                                dbAccount.Attributes.Contains("accountnumber") &&
                                dbAccount.Attributes.Contains("cxp_relatedclientnumber") &&
                                dbAccount["accountnumber"] != null &&
                                dbAccount["cxp_relatedclientnumber"] != null &&
                                dbAccount["accountnumber"] == dbAccount["cxp_relatedclientnumber"])
                            {
                                modifiedOpportunity["cxp_relatedclientnumber"] = dbAccount["cxp_relatedclientnumber"];
                            }

                            objService.Update(modifiedOpportunity);
                        }


                        List<Entity> matters = pluginUtilities.GetAllMattersForAccount(account.Id, objService, objtracer);

                        foreach (Entity currMatter in matters)
                        {
                            Entity modifiedMatter = new Entity("cpdc_matter", currMatter.Id);
                            modifiedMatter["cxp_clientnumber"] = accountNumber;
                            if (dbAccount != null &&
                                dbAccount.Attributes.Contains("accountnumber") &&
                                dbAccount.Attributes.Contains("cxp_relatedclientnumber") &&
                                dbAccount["accountnumber"] != null &&
                                dbAccount["cxp_relatedclientnumber"] != null &&
                                dbAccount["accountnumber"] == dbAccount["cxp_relatedclientnumber"])
                            {
                                modifiedMatter["cxp_relatedclientnumber"] = dbAccount["cxp_relatedclientnumber"];
                            }

                            objService.Update(modifiedMatter);
                        }
                       

                    }


                    //Added: On Update of CR Industry field from Account update the CR Industry field in all related opportunities 
                    if (entity.GetAttributeValue<EntityReference>("cxp_crindustrypracticearea1") != null)
                    {
                        //Guid accountCRindustryGuid = entity.GetAttributeValue<EntityReference>("cxp_crindustrypracticearea1").Id;

                        var industryCR = entity.GetAttributeValue<EntityReference>("cxp_crindustrypracticearea1");


                        List<Entity> relatedOpportunities = pluginUtilities.GetAllOpportunitiesForAccount(account.Id, objService, objtracer);

                        foreach (Entity crOpportunity in relatedOpportunities)
                        {
                            Entity modifiedCROpportunity = new Entity("opportunity", crOpportunity.Id);
                            

                            modifiedCROpportunity["cxp_industrypracticearea"] = industryCR;

                            objService.Update(modifiedCROpportunity);
                        }
                    }

                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in PushNewClientNumberToRelatedOppsAndMatters.", ex);
                }
                //

                                                                         
                catch (Exception ex)
                {
                    objtracer.Trace("PushNewClientNumberToRelatedOppsAndMatters: {0}", ex.ToString());
                    //throw;
                    throw new InvalidPluginExecutionException("An error occurred in PushNewClientNumberToRelatedOppsAndMatters.", ex);
                }
            }
        }
    }
}
