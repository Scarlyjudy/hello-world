using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CXPPlugins
{
    public class ContactOnUpdateTextLookupFieldsMarketo : IPlugin
    {
        //private IOrganizationService service;
        //private ITracingService tracer;
        public bool changeflag = false;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            ITracingService objTracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            objTracer.Trace("tracer created in execute method");

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
                //tracer.Trace("Target is entity");
                Entity postContact = context.PostEntityImages["PostImage"];
                Entity preContact = context.PreEntityImages["PreImage"];
                Guid ContactId = context.PrimaryEntityId;
                Entity modifiedContact = new Entity("contact", ContactId);

                // country/region
                if (postContact != null && postContact.Contains("cxp_marketo_countryregion"))
                {
                    // check that prevents infinite loop with workflow
                    bool flag = checkIfSameName(preContact, postContact, "cxp_marketo_countryregion");

                    if (flag == true)
                    {
                        string name = postContact.GetAttributeValue<string>("cxp_marketo_countryregion");
                        string type = "cpdc_country";
                        string primary_name = "cpdc_name";
                        string guid_field = "cpdc_countryid";
                        Guid recordguid = Guid.Empty;
                        recordguid = GetEntityRecordGuid(name, type, primary_name, guid_field, objService, objTracer);

                        if (recordguid != Guid.Empty)
                        {
                            modifiedContact["cpdc_contactcountryid"] = new EntityReference(type, recordguid);
                            changeflag = true;
                            objTracer.Trace("Successfully staged lookup field: " + name);
                        }
                    }
                    else
                    {
                        objTracer.Trace("Correct lookup field detected for " + "cxp_marketo_countryregion" + ". No changes have been made.");
                    }
                } //end country/region 

                // state/province
                if (postContact != null && postContact.Contains("cxp_marketo_stateprovince"))
                {
                    // check that prevents infinite loop with workflow
                    bool flag = checkIfSameName(preContact, postContact, "cxp_marketo_stateprovince");

                    if (flag == true)
                    {
                        string name = postContact.GetAttributeValue<string>("cxp_marketo_stateprovince");
                        string type = "cpdc_state";
                        string primary_name = "cpdc_name";
                        string guid_field = "cpdc_stateid";
                        Guid recordguid = Guid.Empty;
                        recordguid = GetEntityRecordGuid(name, type, primary_name, guid_field, objService, objTracer);

                        if (recordguid != Guid.Empty)
                        {
                            modifiedContact["cxp_stateprovince"] = new EntityReference(type, recordguid);
                            changeflag = true;
                            objTracer.Trace("Successfully staged lookup field: " + name);
                        }
                    }
                    else
                    {
                        objTracer.Trace("Correct lookup field detected for " + "cxp_marketo_stateprovince" + ". No changes have been made.");
                    }
                } //end state/province

                // Account
                if (postContact != null && postContact.Contains("cxp_marketo_account"))
                {
                    // check that prevents infinite loop with workflow
                    bool flag = checkIfSameName(preContact, postContact, "cxp_marketo_account");

                    if (flag == true)
                    {
                        string name = postContact.GetAttributeValue<string>("cxp_marketo_account");
                        string type = "account";
                        string primary_name = "name";
                        string guid_field = "accountid";
                        Guid recordguid = Guid.Empty;
                        recordguid = GetEntityRecordGuid(name, type, primary_name, guid_field, objService, objTracer);

                        if (recordguid != Guid.Empty)
                        {
                            modifiedContact["parentcustomerid"] = new EntityReference(type, recordguid);
                            changeflag = true;
                            objTracer.Trace("Successfully staged lookup field: " + name);
                        }
                    }
                    else
                    {
                        objTracer.Trace("Correct lookup field detected for " + "cxp_marketo_account" + ". No changes have been made.");
                    }
                } //end Account

                // Record Source
                if (postContact != null && postContact.Contains("cxp_marketo_recordsourceid"))
                {
                    // check that prevents infinite loop with workflow
                    bool flag = checkIfSameName(preContact, postContact, "cxp_marketo_recordsourceid");

                    if (flag == true)
                    {
                        string name = postContact.GetAttributeValue<string>("cxp_marketo_recordsourceid");
                        string type = "campaign";
                        string primary_name = "name";
                        string guid_field = "campaignid";
                        Guid recordguid = Guid.Empty;
                        recordguid = GetEntityRecordGuid(name, type, primary_name, guid_field, objService, objTracer);

                        if (recordguid != Guid.Empty)
                        {
                            modifiedContact["cpdc_recordsourceid"] = new EntityReference(type, recordguid);
                            changeflag = true;
                            objTracer.Trace("Successfully staged lookup field: " + name);
                        }
                    }
                    else
                    {
                        objTracer.Trace("Correct lookup field detected for " + "cxp_marketo_recordsourceid" + ". No changes have been made.");
                    }
                } //end Record Source

                if (changeflag == true)
                {
                    objService.Update(modifiedContact);
                    objTracer.Trace("Lookup Fields updated successfully.");
                }
            }
        }


        //Get record for the specified entity and name
        protected Guid GetEntityRecordGuid(string recordname, string entitytype, string primaryname, string guidfield, IOrganizationService objService, ITracingService objTracer)
        {
            try
            {
                Guid system_type = Guid.Empty;
                //Construct query
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = primaryname;
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(recordname);
                ColumnSet columns = new ColumnSet(guidfield);

                QueryExpression query = new QueryExpression();
                query.ColumnSet = columns;
                query.EntityName = entitytype;
                query.Criteria.AddCondition(condition);
                query.NoLock = true;

                EntityCollection result = objService.RetrieveMultiple(query);
                if (result.Entities.Count() > 0)
                {
                    system_type = result[0].GetAttributeValue<Guid>(guidfield); //always grab the first match
                    objTracer.Trace("Record found. Matched " + entitytype + " record GUID: " + system_type.ToString());
                }
                else
                {
                    objTracer.Trace("No record matches found.");
                }
                return system_type;
            }
            catch (Exception ex)
            {
                objTracer.Trace("LeadOnUpdateTextLookupFieldsMarketo: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("LeadOnUpdateTextLookupFieldsMarketo: {0}" + ex.ToString());
            }
        }

        protected bool checkIfSameName(Entity preImg, Entity postImg, string fieldname)
        {
            bool check = false;
            if (preImg != null && preImg.Contains(fieldname))
            {
                if (postImg.GetAttributeValue<string>(fieldname) != preImg.GetAttributeValue<string>(fieldname))
                {
                    check = true;
                }
            }
            else
            { check = true; }
            return check;
        }
    }
}
