# hello-world
freecode camp start
learning new technology and software
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
//checking

namespace CXPPlugins
{
    public class AccountCompanyTypeCreateDelete : IPlugin
    {
        //private IOrganizationService service;
        //private ITracingService tracer;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            EntityReference targetEntity = null;
            string relationshipName = string.Empty;
            EntityReferenceCollection relatedEntities = null;
            Entity companyType = null;
            Guid companyTypeId = Guid.Empty;
            Guid accountId = Guid.Empty;

            try
            {
                if (context.MessageName.ToLower() == "associate" || context.MessageName.ToLower() == "disassociate")
                {
                    // Get the relationship key from context
                    if (context.InputParameters.Contains("Relationship"))
                    {
                        // get the relationship name for which the plugin fired
                        relationshipName = ((Relationship)context.InputParameters["Relationship"]).SchemaName;
                    }

                    // check if it is the account company type relationship
                    if (relationshipName != "cxp_cxp_companytype_account")
                    {
                        return;
                    }

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                    {
                        targetEntity = (EntityReference)context.InputParameters["Target"];

                        if (context.MessageName.ToLower() == "associate")
                        {
                            if (targetEntity.LogicalName != "cxp_companytype")
                            {
                                return;
                            }
                            else
                            {
                                companyType = service.Retrieve("cxp_companytype", targetEntity.Id, new ColumnSet("cxp_name", "cxp_companytypeid"));
                            }
                        }
                        else if (context.MessageName.ToLower() == "disassociate")
                        {
                            if (targetEntity.LogicalName != "account")
                            {
                                return;
                            }
                            else
                            {
                                if (context.InputParameters.Contains("RelatedEntities") && context.InputParameters["RelatedEntities"] is EntityReferenceCollection)
                                {
                                    relatedEntities = context.InputParameters["RelatedEntities"] as EntityReferenceCollection;

                                    if (relatedEntities != null && relatedEntities.Count() > 0)
                                    {
                                        companyType = service.Retrieve("cxp_companytype", relatedEntities[0].Id, new ColumnSet("cxp_name", "cxp_companytypeid"));
                                    }
                                }
                            }
                        }

                        Entity modifiedAccount = new Entity("account");
                        if (context.MessageName.ToLower() == "associate")
                        {
                            if (context.InputParameters.Contains("RelatedEntities") &&
                            context.InputParameters["RelatedEntities"] is EntityReferenceCollection)
                            {
                                relatedEntities = context.InputParameters["RelatedEntities"] as EntityReferenceCollection;

                                foreach (EntityReference currAccountRef in relatedEntities)
                                {
                                    modifiedAccount = new Entity("account");
                                    modifiedAccount["accountid"] = currAccountRef.Id;
                                    //if (context.MessageName.ToLower() == "associate")
                                    //{
                                    if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "investor")
                                    {
                                        modifiedAccount["cxp_investorcompany"] = true;
                                    }
                                    if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "management company")
                                    {
                                        modifiedAccount["cxp_managementcompany"] = true;
                                    }
                                    if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "private equity firm")
                                    {
                                        modifiedAccount["cxp_privateequityfirm"] = true;
                                    }
                                    //}
                                    /*else if (context.MessageName.ToLower() == "disassociate")
                                    {
                                        if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "investor")
                                        {   
                                            modifiedAccount["cxp_investorcompany"] = false;
                                        }
                                        else if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "management company")
                                        {
                                            modifiedAccount["cxp_managementcompany"] = false;
                                        }
                                        if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "private equity firm")
                                        {
                                            modifiedAccount["cxp_privateequityfirm"] = false;
                                        }
                                    }*/
                                    service.Update(modifiedAccount);
                                }
                            }
                        }
                        else if (context.MessageName.ToLower() == "disassociate")
                        {
                            modifiedAccount["accountid"] = targetEntity.Id;

                            if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "investor")
                            {
                                modifiedAccount["cxp_investorcompany"] = false;
                            }
                            else if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "management company")
                            {
                                modifiedAccount["cxp_managementcompany"] = false;
                            }
                            if (companyType.GetAttributeValue<string>("cxp_name").ToLower() == "private equity firm")
                            {
                                modifiedAccount["cxp_privateequityfirm"] = false;
                            }

                            service.Update(modifiedAccount);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                tracer.Trace("AccountCompanyTypeCreateDelete: error - " + ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException("AccountCompanyTypeCreateDelete: error - " + ex.Message + "\n" + ex.StackTrace);
            }
        }
    }

}
