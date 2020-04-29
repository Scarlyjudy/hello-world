using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CXPPlugins
{
    public class CheckIfGFRClientNameNeeded : IPlugin
    {
        //private Entity postOpportunity;

        //private IOrganizationService service;
        //private ITracingService tracer;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            ITracingService ObjTracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            ObjTracer.Trace("tracer created in execute method");

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService ObjService = serviceFactory.CreateOrganizationService(context.UserId);

            // The InputParameters collection contains all the data passed in the message request.
            if (context.MessageName.ToLower().Equals("delete"))
            {
                Entity preOpportunity = context.PreEntityImages["PreImage"];

                if (preOpportunity != null && preOpportunity.Contains("parentaccountid"))
                {
                    EntityReference oldAccountRef = preOpportunity.GetAttributeValue<EntityReference>("parentaccountid");

                    if (oldAccountRef != null)
                    {
                        // get all opportunityes for the old account
                        List<Entity> opportunities = GetOpportunitiesForAccount(oldAccountRef, ObjService, ObjTracer);
                        ProcessOpportunities(opportunities, oldAccountRef, ObjService, ObjTracer);
                    }
                }
            }
            else
            {
                if (context.InputParameters.Contains("Target") &&
                    context.InputParameters["Target"] is Entity)
                {
                    // Obtain the target entity from the input parameters.
                    Entity entity = (Entity)context.InputParameters["Target"];

                    // Verify that the target entity represents an entity type you are expecting. 
                    // For example, an account. If not, the plug-in was not registered correctly.
                    if (entity.LogicalName != "opportunity")
                    {
                        ObjTracer.Trace("entity is not an opportunity");
                        return;
                    }

                    try
                    {
                        Entity preOpportunity = null;

                        if (!context.MessageName.ToLower().Equals("create"))
                        {
                            // Get the before image
                            preOpportunity = context.PreEntityImages["PreImage"];
                        }

                        // Get the opportunity id from the opportunity entity
                        Guid opportunityId = entity.Id;
                        EntityReference newAccountRef = null;
                        EntityReference oldAccountRef = null;

                        // get the new and old parent account references
                        if (context.MessageName.ToLower().Equals("create"))
                        {
                            if (entity.Contains("parentaccountid"))
                            {
                                newAccountRef = entity.GetAttributeValue<EntityReference>("parentaccountid");
                            }
                        }
                        else if (context.MessageName.ToLower().Equals("update"))
                        {

                            Entity postOpportunity = context.PostEntityImages["PostImage"];
                            if (entity.Contains("parentaccountid"))
                            {
                                newAccountRef = entity.GetAttributeValue<EntityReference>("parentaccountid");
                            }
                            else
                            {
                                newAccountRef = GetAccountForOpportunity(opportunityId, ObjService, ObjTracer, postOpportunity);
                            }

                            if (preOpportunity.Contains("parentaccountid"))
                            {
                                oldAccountRef = preOpportunity.GetAttributeValue<EntityReference>("parentaccountid");
                            }
                        }

                        List<Entity> opportunities = null;

                        if (newAccountRef != null)
                        {
                            // get all opportunities for the new account
                            opportunities = GetOpportunitiesForAccount(newAccountRef, ObjService, ObjTracer);
                            ProcessOpportunities(opportunities, newAccountRef, ObjService, ObjTracer);
                        }

                        if (oldAccountRef != null)
                        {
                            // get all opportunityes for the old account
                            opportunities = GetOpportunitiesForAccount(oldAccountRef, ObjService, ObjTracer);
                            ProcessOpportunities(opportunities, oldAccountRef, ObjService, ObjTracer);
                        }
                        //throw new Exception("debugger");
                    }

                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        ObjTracer.Trace("CheckIfGFRClientNameNeeded: {0}", ex.Message + "\n" + ex.StackTrace);
                        throw new InvalidPluginExecutionException("An error occurred in CheckIfGFRClientNameNeeded.", ex);
                    }

                    catch (Exception ex)
                    {
                        ObjTracer.Trace("CheckIfGFRClientNameNeeded: {0}", ex.Message + "\n" + ex.StackTrace);
                        throw new InvalidPluginExecutionException("CheckIfGFRClientNameNeeded: {0}" + ex.Message + "\n" + ex.StackTrace);
                    }
                }
            }
        }

        protected EntityReference GetAccountForOpportunity(Guid opportunityId, IOrganizationService ObjService, ITracingService ObjTracer, Entity postOpportunity)
        {
            EntityReference accountRef = null;

            try
            {
                //Entity opportunity = ObjService.Retrieve("opportunity", opportunityId, new ColumnSet(new string[] { "parentaccountid"}));
                Entity opportunity = postOpportunity;
                if (opportunity != null)
                {
                    accountRef = opportunity.GetAttributeValue<EntityReference>("parentaccountid");
                }
            }
            catch (Exception ex)
            {
                ObjTracer.Trace("CheckIfGFRClientNameNeeded: {0}", ex.Message + "\n" + ex.StackTrace);
                throw;
            }

            return accountRef;
        }

        protected bool IsGFRRequired(EntityReference serviceAreaRef, IOrganizationService ObjService, ITracingService ObjTracer)
        {
            bool bReturnValue = false;

            try
            {
                Entity serviceArea = ObjService.Retrieve("rg_servicearea", serviceAreaRef.Id, new ColumnSet(new string[] { "cxp_gfrneeded" }));
                if (serviceArea != null)
                {
                    bReturnValue = serviceArea.GetAttributeValue<bool>("cxp_gfrneeded");
                }
            }
            catch (Exception ex)
            {
                ObjTracer.Trace("CheckIfGFRClientNameNeeded: {0}", ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException("CheckIfGFRClientNameNeeded: {0}" + ex.Message + "\n" + ex.StackTrace);

            }
            return bReturnValue;
        }

        protected List<Entity> GetOpportunitiesForAccount(EntityReference accountRef, IOrganizationService ObjService, ITracingService ObjTracer)
        {
            List<Entity> opportunities = new List<Entity>();

            try
            {
                QueryByAttribute query = new QueryByAttribute("opportunity");
                query.ColumnSet = new ColumnSet(new string[] { "cxp_servicerequested", "parentaccountid", "opportunityid" });
                query.AddAttributeValue("parentaccountid", accountRef.Id);

                EntityCollection results = ObjService.RetrieveMultiple(query);

                if (results != null)
                {
                    opportunities = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                ObjTracer.Trace("CheckIfGFRClientNameNeeded: {0}", ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException("CheckIfGFRClientNameNeeded: {0}" + ex.Message + "\n" + ex.StackTrace);

            }

            return opportunities;
        }

        protected void UpdateAccount(EntityReference accountRef, bool bGFRRequired, IOrganizationService ObjService, ITracingService ObjTracer)
        {
            try
            {
                Entity modifiedAccount = new Entity("account");
                modifiedAccount["accountid"] = accountRef.Id;
                modifiedAccount["cxp_gfrrequired"] = bGFRRequired;

                ObjService.Update(modifiedAccount);
            }
            catch (Exception ex)
            {
                ObjTracer.Trace("CheckIfGFRClientNameNeeded: {0}", ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException("CheckIfGFRClientNameNeeded: {0}" + ex.Message + "\n" + ex.StackTrace);
            }
        }

        protected void ProcessOpportunities(List<Entity> opportunities, EntityReference accountRef, IOrganizationService ObjService, ITracingService ObjTracer)
        {
            bool bGFRRequired = false;

            foreach (Entity currOpportunity in opportunities)
            {
                EntityReference serviceAreaRef = currOpportunity.GetAttributeValue<EntityReference>("cxp_servicerequested");
                if (serviceAreaRef != null)
                {
                    bGFRRequired = IsGFRRequired(serviceAreaRef, ObjService, ObjTracer);
                    if (bGFRRequired)
                    {
                        break;
                    }
                }
            }

            // set the GFR Required flag on the account
            UpdateAccount(accountRef, bGFRRequired, ObjService, ObjTracer);
        }
    }
}
 
