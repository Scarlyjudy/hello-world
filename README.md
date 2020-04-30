using System;
using Microsoft.Xrm.Sdk;

//The plugin is called on create of new Opportunnity and update of business development fields in opportunity entity
namespace CXPPlugins
{
    public class CreateOppExtension : IPlugin
    {
        //private IOrganizationService service;
        //private ITracingService tracer;
        public bool changeflag = false;

        //TODO: Take a look at the images that you have on plugins. Only the attributes that you are using need to be passed in.
        //TODO: I suggest that you update values that you use the post entity to update values on CR extension that aren't actually the same. This will allow you to recover from an error on an earlier update.
        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            //tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            //tracer.Trace("tracer created in execute method");

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.UserId);

            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            #region Early Return Checks
            if (context.Depth > 3)
            {
                return;
            }
            //If we dont' have an target or target isn't entity or was triggered from delete return
            if (!context.InputParameters.Contains("Target")
                || !(context.InputParameters["Target"] is Entity)
                || context.MessageName.Equals("delete", StringComparison.OrdinalIgnoreCase))
            {
                tracer.Trace("Trigger was not correct, or the target wasn't an entity.");
                return;
            }
            Entity objOpportunity = (Entity)context.InputParameters["Target"];
            //If the target is null or it's not an opportunity return
            if (objOpportunity == null ||
                !objOpportunity.LogicalName.Equals("opportunity", StringComparison.OrdinalIgnoreCase))
            {
                tracer.Trace("The correct entity was not passed into the plugin");
                return;
            }
            //If we don't have a parent or the parent we have is null then return
            ////Using post-image because that is the only way we guarantee that it is passed if there is one.
            //if (!context.PostEntityImages["PostImage"].Contains("cxp_parentopportunity")
            //    || context.PostEntityImages["PostImage"]["cxp_parentopportunity"] == null)
            //{
            //    tracer.Trace("The opportunity has no parent, so we do not need execute the logic for this plugin.");
            //    return;
            //}
            #endregion

            //TODO: Could no longer follow link.
            //https://docs.microsoft.com/en-us/previous-versions/dynamicscrm-2016/developers-guide/gg326129(v%3Dcrm.8)
            Guid guOppID = context.PrimaryEntityId; //TODO: Do you need this?

            //TODO: Read below comment
            //Check here if post image has a reference to this entity if it does then set the GUID here otherwise set just the name.
            //This will allow you to write your code so that if something happens and it doesn't create one on create of the entity,
            //than you will be able to create one even on updates
            //Also right now you are creating if for some reason you have a null value for objOpportunity
            Entity objOppExtension = new Entity("cxp_oppextensionforcresalesvalidation");
            if (objOpportunity != null)
            {
                //TODO: Seeing if a entity contains attribute: objOpportunity.Contains("name") seeing if attribute is null: objOpportunity["name"] == null
                if (objOpportunity.Attributes.ContainsKey("name") && objOpportunity.Attributes["name"] != null)
                {
                    //TODO: objOpportunity.GetAttributeValue<string>("name"); Can be used everywhere in this... optionsetValue, Money, decimal, string, EntityReference, string
                    //https://community.dynamics.com/crm/b/crmentropy/archive/2013/08/28/entity-getattributevalue-lt-t-gt-explained
                    objOppExtension["cxp_name"] = objOpportunity["name"].ToString();
                }

                if (objOpportunity.Attributes.ContainsKey("rg_gsmc") && objOpportunity.Attributes["rg_gsmc"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_businessdevelopment"] = Convert.ToBoolean(objOpportunity["rg_gsmc"].ToString());
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cpdc_targetstatus") && objOpportunity.Attributes["cpdc_targetstatus"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_targetstatus"] = new OptionSetValue(Convert.ToInt32(((OptionSetValue)objOpportunity["cpdc_targetstatus"]).Value));
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cpdc_assignedcre") && objOpportunity.Attributes["cpdc_assignedcre"] != null)
                {

                    //TODO: Is this already an entity reference to system user? if so objOpportunity.GetAttributeValue<EntityReference>("cpdc_assignedcre");
                    EntityReference objAssignedCRE = ((EntityReference)objOpportunity["cpdc_assignedcre"]);
                    objOppExtension["cxp_assignedcre"] = new EntityReference("systemuser", objAssignedCRE.Id);
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cpdc_riskweightedvalue") && objOpportunity.Attributes["cpdc_riskweightedvalue"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_riskweightedvalue"] = new Money(Convert.ToDecimal(objOpportunity["cpdc_riskweightedvalue"].ToString()));
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cxp_neworexistingclient") && objOpportunity.Attributes["cxp_neworexistingclient"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_neworexistingclient"] = new OptionSetValue(Convert.ToInt32(((OptionSetValue)objOpportunity["cxp_neworexistingclient"]).Value));
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cxp_typeofwork") && objOpportunity.Attributes["cxp_typeofwork"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_typeofwork"] = new OptionSetValue(Convert.ToInt32(((OptionSetValue)objOpportunity["cxp_typeofwork"]).Value));
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cxp_validationsought") && objOpportunity.Attributes["cxp_validationsought"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_validationsought"] = new OptionSetValue(Convert.ToInt32(((OptionSetValue)objOpportunity["cxp_validationsought"]).Value));
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cxp_describehowthisopportunitywasidentified") && objOpportunity.Attributes["cxp_describehowthisopportunitywasidentified"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_describehowthisopportunitywasidentified"] = objOpportunity["cxp_describehowthisopportunitywasidentified"].ToString();
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cxp_whatworkdidyouprovidefromidentificationto") && objOpportunity.Attributes["cxp_whatworkdidyouprovidefromidentificationto"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_whatworkdidyouprovidefromidentificationto"] = objOpportunity["cxp_whatworkdidyouprovidefromidentificationto"].ToString();
                }
                //
                if (objOpportunity.Attributes.ContainsKey("cxp_describeyourrelationshiptotheclientthatha") && objOpportunity.Attributes["cxp_describeyourrelationshiptotheclientthatha"] != null)
                {
                    //TODO: GetAttributeValue
                    objOppExtension["cxp_describeyourrelationshiptotheclientthatha"] = objOpportunity["cxp_describeyourrelationshiptotheclientthatha"].ToString();
                }

                // Set LookUp value for Opp On CRE Extension                                     

                    objOppExtension["cxp_opportunity"] = new EntityReference("opportunity", guOppID);
                


            }
            //Retrieve the record's Guid based upon passed in parameters.
            if (context.MessageName.ToLower() == "create")
            {
                Guid guOppExtensionID = objService.Create(objOppExtension);

                if (guOppExtensionID != Guid.Empty)
                {
                    Entity objModifiedOpp = new Entity("opportunity");
                    objModifiedOpp.Id = objOpportunity.Id;
                    objModifiedOpp["cxp_opportunityextension"] = new EntityReference("cxp_oppextensionforcresalesvalidation", guOppExtensionID);
                    objService.Update(objModifiedOpp);



                }



            }
            //Registered the Post image to retrieve the opp extention data if that is not updated in the particular entity.
            else if (context.MessageName.ToLower() == "update")
            {
                Entity postOpportunity = context.PostEntityImages["PostImage"];
                if (postOpportunity != null && postOpportunity.Attributes.ContainsKey("cxp_opportunityextension"))
                {
                    EntityReference objOppExt = (EntityReference)postOpportunity.Attributes["cxp_opportunityextension"];
                    objOppExtension.Id = objOppExt.Id;

                    try
                    {
                        objService.Update(objOppExtension);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidPluginExecutionException(" exception " + ex.Message + ex.StackTrace);
                    }
                }
            }
        }
    }
}
