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

namespace CXPPlugins
{
    public class AddUserToAccessTeamForFollow : IPlugin
    {
        //private IOrganizationService service;
        //private ITracingService tracer;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            ITracingService objTracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.UserId);

            // The InputParameters collection contains all the data passed in the message request.
            if (context.MessageName.Equals("Delete"))
            {
                Entity preFollow = context.PreEntityImages["PreImage"];

                if (preFollow != null)
                {
                    if (preFollow.Contains("ownerid") && preFollow.Contains("regardingobjectid"))
                    {
                        Guid userId = preFollow.GetAttributeValue<EntityReference>("ownerid").Id;
                        EntityReference record = preFollow.GetAttributeValue<EntityReference>("regardingobjectid");
                        string recordType = record.LogicalName;
                        Guid recordId = record.Id;

                        //EA in Account Check
                        if (recordType.ToLower().Equals("account"))
                        {
                            bool updated = updateRecordAssociatedEAField(userId, recordId, "account", "accountid", "delete", objService, objTracer);
                            if (updated) { objTracer.Trace("Associated EAs field updated (removed) successfully for Account."); }
                        }

                        //EA in Opportunity Check
                        if (recordType.ToLower().Equals("opportunity"))
                        {
                            bool updated = updateRecordAssociatedEAField(userId, recordId, "opportunity", "opportunityid", "delete", objService, objTracer);
                            if (updated) { objTracer.Trace("Associated EAs field updated (removed) successfully for Opportunity."); }
                        }
                    }
                }
            } //END DELETE

            else
            {
                if (context.InputParameters.Contains("Target") &&
                    context.InputParameters["Target"] is Entity)
                {
                    // Obtain the target entity from the input parameters.
                    Entity entity = (Entity)context.InputParameters["Target"];

                    // Verify that the target entity represents an entity type you are expecting. 
                    // For example, an account. If not, the plug-in was not registered correctly.
                    //if (entity.LogicalName != "postfollow")
                    //{
                    //    tracer.Trace("entity is not postfollow");
                    //    return;
                    //}

                    if (entity.LogicalName != "cxp_cxpfollow")
                    {
                        objTracer.Trace("Entity is not CXP Follow");
                        return;
                    }

                    try
                    {
                        if (context.MessageName.Equals("Create"))
                        {
                            Guid userId = entity.GetAttributeValue<EntityReference>("ownerid").Id;
                            EntityReference record = entity.GetAttributeValue<EntityReference>("regardingobjectid");
                            string recordType = record.LogicalName;
                            Guid recordId = record.Id;

                            if (recordType.ToLower().Equals("account") || recordType.ToLower().Equals("contact"))
                            {
                                PluginUtilities pluginUtilities = new PluginUtilities();
                                //pluginUtilities.Service = objService;
                                //pluginUtilities.Tracer = objTracer;
                                //cxpUtilities.CheckIfUserIsAlreadyOnAccessTeam(userId, recordType, recordId);
                                pluginUtilities.AddUserToRecordAccessTeam(userId, recordType, recordId, objService, objTracer);
                            }

                            //EA in Account Check
                            if (recordType.ToLower().Equals("account"))
                            {
                                bool updated = updateRecordAssociatedEAField(userId, recordId, "account", "accountid", "insert", objService, objTracer);
                                if (updated) { objTracer.Trace("Associated EAs field updated (insert) successfully for Account."); }
                            }

                            //EA in Opportunity Check
                            if (recordType.ToLower().Equals("opportunity"))
                            {
                                bool updated = updateRecordAssociatedEAField(userId, recordId, "opportunity", "opportunityid", "insert", objService, objTracer);
                                if (updated) { objTracer.Trace("Associated EAs field updated (insert) successfully for Opportunity."); }
                            }
                        } //END CREATE
                    }

                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        throw new InvalidPluginExecutionException("An error occurred in MyPlug-in.", ex);
                    }

                    catch (Exception ex)
                    {
                        objTracer.Trace("MyPlugin: {0}", ex.ToString());
                        //throw ex;
                        throw new InvalidPluginExecutionException("MyPlugin: " + ex.ToString());
                    }
                }
            }
        }

        private bool updateRecordAssociatedEAField(Guid userId, Guid recordId, string recordtype, string idfieldname, string action, IOrganizationService objService, ITracingService objTracer)
        {
            bool updated = false;
            //check if the user has an executive assistant
            PluginUtilities pluginUtilities = new PluginUtilities();
            //pluginUtilities.Service = service;
            //pluginUtilities.Tracer = tracer;
            EntityReference exec_assistant = pluginUtilities.GetUserEA(userId, objService, objTracer);
            string updated_eas_str = null;
            string original_ea_str = null;
            bool needToUpdate = false;

            if (exec_assistant != null)
            {
                if (action.Equals("insert"))
                {
                    Entity tmpEntity = objService.Retrieve(recordtype, recordId, new ColumnSet(new string[] { "cxp_associatedeas" }));

                    if (tmpEntity != null)
                    {
                        original_ea_str = tmpEntity.GetAttributeValue<string>("cxp_associatedeas");
                    }

                    if (original_ea_str == null)
                    {
                        updated_eas_str = exec_assistant.Name.ToString();
                        needToUpdate = true;

                    }
                    else if (original_ea_str.Contains(exec_assistant.Name.ToString()) == false)
                    {
                        updated_eas_str = original_ea_str + "; " + exec_assistant.Name.ToString();
                        needToUpdate = true;
                    }

                    if (needToUpdate)
                    {
                        Entity modifiedEntity = new Entity(recordtype);
                        modifiedEntity[idfieldname] = recordId;
                        modifiedEntity["cxp_associatedeas"] = updated_eas_str;
                        objService.Update(modifiedEntity);

                        updated = true;
                    }
                } //end insert

                else if (action.Equals("delete"))
                {
                    Entity tmpEntity = objService.Retrieve(recordtype, recordId, new ColumnSet(new string[] { "cxp_associatedeas" }));

                    if (tmpEntity != null)
                    {
                        original_ea_str = tmpEntity.GetAttributeValue<string>("cxp_associatedeas");
                    }

                    if (original_ea_str != null)
                    {
                        if (original_ea_str.Contains("; " + exec_assistant.Name.ToString()))
                        {
                            updated_eas_str = original_ea_str.Replace("; " + exec_assistant.Name.ToString(), "");
                            needToUpdate = true;
                        }

                        else if (original_ea_str.Contains(exec_assistant.Name.ToString()))
                        {
                            updated_eas_str = original_ea_str.Replace(exec_assistant.Name.ToString(), "");
                            //format string if needed
                            if (updated_eas_str.Contains("; ;"))
                            {
                                updated_eas_str = updated_eas_str.Replace("; ;", ";");
                            }

                            needToUpdate = true;
                        }
                    }
                    if (needToUpdate)
                    {
                        Entity modifiedEntity = new Entity(recordtype);
                        modifiedEntity[idfieldname] = recordId;
                        modifiedEntity["cxp_associatedeas"] = updated_eas_str;
                        objService.Update(modifiedEntity);

                        updated = true;
                    }
                } //end delete
            }
            return updated;
        }
    }
}
