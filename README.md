using System;
using Microsoft.Xrm.Sdk;

//The plugin is called on create of new account and update of Client Segment and SAM Status fields in account entity
namespace CXPPlugins
{
    public class CreateAccountExtension : IPlugin
    {

        public bool changeflag = false;

        public void Execute(IServiceProvider serviceProvider)
        {
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
            if (context.Depth > 1)
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
            Entity objAccount = (Entity)context.InputParameters["Target"];
            //If the target is null or it's not an Account return
            if (objAccount == null ||
                !objAccount.LogicalName.Equals("account", StringComparison.OrdinalIgnoreCase))
            {
                tracer.Trace("The correct entity was not passed into the plugin");
                return;
            }

            Entity objAccPreImage = context.PreEntityImages != null && context.PreEntityImages.Count > 0 ? context.PreEntityImages["PreImage"] : null;
            Entity objAccPostImage = context.PostEntityImages != null && context.PostEntityImages.Count > 0 ? context.PostEntityImages["PostImage"] : null;
            //If we don't have Client segment or sam status in the image then return
            //if (!context.PreEntityImages["PreImage"].Contains("cxp_clientsegment")
            //    || context.PreEntityImages["PreImage"]["cxp_clientsegment"] == null
            //    || !context.PreEntityImages["PreImage"].Contains("cxp_samstatus")
            //    || context.PreEntityImages["PreImage"]["cxp_samstatus"] == null)
            //{
            //    tracer.Trace("The account doesn't have relevent image fields.");
            //    return;
            //}
            #endregion

            Guid guAccID = context.PrimaryEntityId; //TODO: Do you need this?
            bool blnCreateFirst = false;
            bool blnCreateSecond = false;

            Entity objAccExtension = new Entity("cxp_accountextension");
            Entity objAccExtension2 = new Entity("cxp_accountextension");
            if (objAccount != null)
            {
                string strOldValue = string.Empty;
                string strNewValue = string.Empty;

                string strOldValue2 = string.Empty;
                string strNewValue2 = string.Empty;

                //As part of Update if we get client segment value in the context Account entity then status change field will be set as client segment
                if (objAccount.Attributes.ContainsKey("cxp_clientsegment") && objAccount.Attributes["cxp_clientsegment"] != null)
                {

                    strOldValue = objAccPreImage != null && objAccPreImage.Attributes.ContainsKey("cxp_clientsegment") && objAccPreImage.Attributes["cxp_clientsegment"] != null ? objAccPreImage.FormattedValues["cxp_clientsegment"].ToString() : string.Empty;
                    strNewValue = objAccPostImage != null && objAccPostImage.Attributes.ContainsKey("cxp_clientsegment") && objAccPostImage.Attributes["cxp_clientsegment"] != null ? objAccPostImage.FormattedValues["cxp_clientsegment"].ToString() : string.Empty;

                    objAccExtension["cxp_statuschangefield"] = new OptionSetValue(280410000);
                    blnCreateFirst = true;
                }

                //As part of Update if we get SAM Status value in the context Account entity then status change field will be set as SAM Status
                if (objAccount.Attributes.ContainsKey("cxp_samstatus") && objAccount.Attributes["cxp_samstatus"] != null)
                {
                    strOldValue2 = objAccPreImage != null && objAccPreImage.Attributes.ContainsKey("cxp_samstatus") && objAccPreImage.Attributes["cxp_samstatus"] != null ? objAccPreImage.FormattedValues["cxp_samstatus"].ToString() : string.Empty;
                    strNewValue2 = objAccPostImage != null && objAccPostImage.Attributes.ContainsKey("cxp_samstatus") && objAccPostImage.Attributes["cxp_samstatus"] != null ? objAccPostImage.FormattedValues["cxp_samstatus"].ToString() : string.Empty;

                    objAccExtension2["cxp_statuschangefield"] = new OptionSetValue(280410001);
                    blnCreateSecond = true;
                }

                //Set old value
                if (!string.IsNullOrEmpty(strOldValue) || !string.IsNullOrEmpty(strOldValue2))
                {
                    objAccExtension["cxp_oldvalue"] = strOldValue;
                    objAccExtension2["cxp_oldvalue"] = strOldValue2;
                }

                //Set new Value
                if (!string.IsNullOrEmpty(strNewValue) || !string.IsNullOrEmpty(strNewValue2))
                {
                    objAccExtension["cxp_newvalue"] = strNewValue;
                    objAccExtension2["cxp_newvalue"] = strNewValue2;
                }

                // Set Date changed field
                objAccExtension["cxp_datechanged"] = DateTime.Now;
                objAccExtension2["cxp_datechanged"] = DateTime.Now;

                // Set Changed by lookup(user))
                EntityReference objModifiedby = ((EntityReference)objAccount["modifiedby"]);
                objAccExtension["cxp_changedby"] = new EntityReference("systemuser", objModifiedby.Id);
                objAccExtension2["cxp_changedby"] = new EntityReference("systemuser", objModifiedby.Id);

                //Set Account lookup
                objAccExtension["cxp_accountname"] = new EntityReference("account", guAccID);
                objAccExtension2["cxp_accountname"] = new EntityReference("account", guAccID);

                // set client segment exception reason
                if (objAccount.Attributes.ContainsKey("cxp_clientsegexceptionreason") && objAccount.Attributes["cxp_clientsegexceptionreason"] != null)
                {
                    objAccExtension["cxp_clientsegexceptionreason"] = objAccount["cxp_clientsegexceptionreason"].ToString();
                }

                ////Set cxp_fiscalyear
                //DateTime now = DateTime.Now;
                //string strCurrentYear = now.ToString("yy");
                //int intCurrentMonth = now.Month;
                //int intActualYear = 0;
                //intActualYear = Convert.ToInt32(strCurrentYear);
                //if (intCurrentMonth != 1)
                //{ //If month not equals to january then adding an year to show next fiscal year else it shows current fiscal year.  
                //    intActualYear = intActualYear + 1;
                //}
                //string strFisYear = "FY" + intActualYear.ToString();

                //objAccExtension["cxp_fiscalyear"] = strFisYear;

            }

            if (blnCreateFirst)
            {
                Guid guAccExtensionID = objService.Create(objAccExtension);
            }
            if (blnCreateSecond)
            {
                Guid guAccExtensionID2 = objService.Create(objAccExtension2);
            }
        }
    }
}
