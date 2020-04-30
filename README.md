using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace CXPWorkflows
{
    public class ReassignLeads : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext context = exusing System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
//using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
namespace CXPWorkflows
{
    public class WorkflowUtilities
    {
        //private IOrganizationService service;
        //private ITracingService tracer;

        //public IOrganizationService Service
        //{
        //    get { return service; }
        //    set { service = value; }
        //}

        //public ITracingService Tracer
        //{
        //    get { return tracer; }
        //    set { tracer = value; }
        //}

        public Entity GetAccount(Guid accountId, ColumnSet columns, IOrganizationService objService, ITracingService objTracer)
        {
            Entity account = null;

            try
            {
                account = objService.Retrieve("account", accountId, columns);
                objTracer.Trace("GetAccount: account = " + account != null ? account.ToString() : "null.");
            }
            catch (Exception ex)
            {
                objTracer.Trace("wfUtilities.GetAccount: " + ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return account;
        }

        public Entity GetContact(Guid contactId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity contact = null;

            try
            {
                contact = objService.Retrieve("contact", contactId, new ColumnSet(new string[] { "createdby" }));
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return contact;
        }

        public Entity GetUser(Guid userId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity user = null;

            try
            {
                user = objService.Retrieve("systemuser", userId, new ColumnSet(new string[] { "businessunitid", "fullname" }));
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return user;
        }

        public string GetSystemParameter(string parameterName, IOrganizationService objService, ITracingService objTracer)
        {
            string returnValue = string.Empty;

            try
            {
                QueryByAttribute query = new QueryByAttribute("cxp_systemparameter");
                query.ColumnSet = new ColumnSet(new string[] { "cxp_value" });
                query.AddAttributeValue("cxp_name", parameterName);

                EntityCollection results = objService.RetrieveMultiple(query);

                Entity currSystemParameter = results.Entities.ElementAt(0);

                returnValue = currSystemParameter.GetAttributeValue<string>("cxp_value");
            }
            catch (Exception ex)
            {
                objTracer.Trace("GetSystemParameter: {0}", ex.ToString());
                //throw;
                throw new InvalidPluginExecutionException(ex.ToString());
            }

            return returnValue;
        }

        public EntityReference GetBUOwnerRef(string buName, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference buOwnerRef = null;

            try
            {
                string systemParameterName = buName + " User";
                QueryByAttribute query = new QueryByAttribute("cxp_systemparameter");
                query.AddAttributeValue("cxp_name", systemParameterName);
                query.ColumnSet = new ColumnSet(new string[] { "cxp_value" });

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null && results.Entities != null && results.Entities.Count() > 0)
                {
                    if (results.Entities[0].Attributes.ContainsKey("cxp_value"))
                    {
                        string buUserName = results.Entities[0].GetAttributeValue<string>("cxp_value");
                        Entity buUser = GetUserByName(buUserName, objService, objTracer);

                        if (buUser != null && buUser.Attributes.ContainsKey("fullname") && buUser.Attributes.ContainsKey("systemuserid"))
                        {
                            buOwnerRef = new EntityReference("systemuser", buUser.GetAttributeValue<Guid>("systemuserid"));
                            buOwnerRef.Name = buUser.GetAttributeValue<string>("fullname");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return buOwnerRef;
        }

        public Entity GetUserByName(string name, IOrganizationService objService, ITracingService objTracer)
        {
            Entity user = null;

            try
            {
                QueryByAttribute query = new QueryByAttribute("systemuser");
                query.AddAttributeValue("fullname", name);
                query.ColumnSet = new ColumnSet(new string[] { "systemuserid", "fullname" });

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null && results.Entities != null && results.Entities.Count() > 0)
                {
                    user = results.Entities[0];
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return user;
        }

        public void AddUserToRecordAccessTeam(Guid userId, string recordType, Guid recordId, IOrganizationService objService, ITracingService objTracer)
        {
            try
            {
                AddUserToRecordTeamRequest request = new AddUserToRecordTeamRequest();
                request.SystemUserId = userId;
                request.Record = new EntityReference(recordType, recordId);
                request.TeamTemplateId = GetAccessTeamTemplateIdByName(recordType, objService, objTracer);

                objService.Execute(request);
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
        }

        public Guid GetAccessTeamTemplateIdByName(String recordType, IOrganizationService objService, ITracingService objTracer)
        {
            Guid templateId = Guid.Empty;
            string accessTeamTemplateName = GetAccessTeamTemplateName(recordType, objService, objTracer);

            if (accessTeamTemplateName != String.Empty)
            {
                try
                {
                    QueryExpression teamQuery = new QueryExpression
                    {
                        NoLock = true,
                        EntityName = "teamtemplate",
                        ColumnSet = new ColumnSet("teamtemplateid"),
                        Criteria =
                        {
                            Conditions =
                            {
                                new ConditionExpression("teamtemplatename", ConditionOperator.Equal, accessTeamTemplateName)
                            }
                        }
                    };

                    var results = objService.RetrieveMultiple(teamQuery).Entities;

                    if (results.Count != 1)
                        //throw new Exception("Can't retrieve Team Access Template for " + recordType);
                        throw new InvalidPluginExecutionException("Can't retrieve Team Access Template for " + recordType);
                    else
                    {
                        templateId = results[0].Id;
                    }
                }
                catch (Exception ex)
                {
                    objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                    //throw ex;
                    throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
                }
            }

            return templateId;
        }

        public string GetAccessTeamTemplateName(string recordType, IOrganizationService objService, ITracingService objTracer)
        {
            string teamName = String.Empty;
            string systemParameterName = String.Empty;

            if (recordType.ToLower().Equals("account"))
            {
                systemParameterName = "Account Access Team Template";
            }
            else if (recordType.ToLower().Equals("contact"))
            {
                systemParameterName = "Contact Access Team Template";
            }

            if (systemParameterName != String.Empty)
            {
                try
                {
                    QueryByAttribute query = new QueryByAttribute("cxp_systemparameter");
                    query.AddAttributeValue("cxp_name", systemParameterName);
                    query.ColumnSet = new ColumnSet(new string[] { "cxp_value" });

                    EntityCollection results = objService.RetrieveMultiple(query);

                    if (results != null)
                    {
                        teamName = results.Entities[0].GetAttributeValue<string>("cxp_value");
                    }
                }
                catch (Exception ex)
                {
                    objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                    //throw ex;
                    throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
                }
            }

            return teamName;
        }

        public void CreateFollowOfRecordForCreatedByUser(Guid userId, string recordType, Guid recordId, IOrganizationService objService, ITracingService objTracer)
        {
            try
            {
                if (UserIsEnabled(userId, objService, objTracer))
                {
                    Entity postFollow = new Entity("postfollow");
                    postFollow.Attributes["regardingobjectid"] = new EntityReference(recordType, recordId);
                    postFollow.Attributes["ownerid"] = new EntityReference("systemuser", userId);
                    objService.Create(postFollow);
                }

            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
        }

        public void CreateCXPFollowOfRecordForCreatedByUser(EntityReference userRef, string recordType, EntityReference recordRef, IOrganizationService objService, ITracingService objTracer)
        {
            try
            {
                if (UserIsEnabled(userRef.Id, objService, objTracer) &&
                    userRef.Name != null &&
                    userRef.Name.ToLower().Contains("admin") == false &&
                    GetCXPFollow(userRef.Id, recordType, recordRef.Id, objService, objTracer) == null)
                {
                    string userName = GetUserFirstNameLastName(userRef, objService, objTracer);
                    Entity cxpFollow = new Entity("cxp_cxpfollow");
                    cxpFollow.Attributes["regardingobjectid"] = new EntityReference(recordType, recordRef.Id);
                    cxpFollow.Attributes["ownerid"] = new EntityReference("systemuser", userRef.Id);
                    cxpFollow.Attributes["subject"] = userName + " follows " + recordRef.Name;
                    objService.Create(cxpFollow);
                }

            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                // throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
        }

        public Entity GetCXPFollow(Guid userId, string recordType, Guid recordId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity cxpFollow = null;

            try
            {
                QueryByAttribute query = new QueryByAttribute("cxp_cxpfollow");
                query.ColumnSet = new ColumnSet("activityid");
                query.AddAttributeValue("ownerid", userId);
                query.AddAttributeValue("regardingobjectid", recordId);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null && results.Entities.Count > 0)
                {
                    cxpFollow = results.Entities[0];
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return cxpFollow;
        }

        public Entity GetAccountCommunicationOptions(Guid accountId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity account = null;

            try
            {
                account = objService.Retrieve("account", accountId, new ColumnSet(new string[] { "donotbulkemail", "donotbulkpostalmail", "donotemail", "donotphone", "donotfax", "donotpostalmail", "donotsendmm" }));
                objTracer.Trace("GetAccount: account = " + account != null ? account.ToString() : "null.");
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return account;
        }

        public Entity GetAccountRFPSECFlags(Guid accountId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity account = null;

            try
            {
                account = objService.Retrieve("account", accountId, new ColumnSet(new string[] { "cxp_rfpquietperiod", "cxp_secblackout", "cxp_rfpquietperiodenddate", "accountid" }));
                objTracer.Trace("GetAccount: account = " + account != null ? account.ToString() : "null.");
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return account;
        }

        public List<Entity> GetContactsForAccount(Guid accountId, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> contacts = new List<Entity>();

            try
            {
                QueryByAttribute query = new QueryByAttribute("contact");
                query.AddAttributeValue("parentcustomerid", accountId);
                query.ColumnSet = new ColumnSet(new string[] { "contactid" });

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    contacts = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return contacts;
        }

        public List<Entity> GetOpportunitiesForAccount(Guid accountId, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> opportunities = new List<Entity>();

            try
            {
                QueryByAttribute query = new QueryByAttribute("opportunity");
                query.AddAttributeValue("parentaccountid", accountId);
                query.ColumnSet = new ColumnSet(new string[] { "opportunityid" });

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    opportunities = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return opportunities;
        }

        public bool boolCheckIfUserIsAlreadyOnAccessTeam(Guid userId, string recordType, Guid recordId, IOrganizationService objService, ITracingService objTracer)
        {
            bool returnValue = false;

            return returnValue;
        }

        public bool UserIsEnabled(Guid userID, IOrganizationService objService, ITracingService objTracer)
        {
            bool bReturnValue = false;

            try
            {
                Entity user = objService.Retrieve("systemuser", userID, new ColumnSet(new string[] { "isdisabled" }));
                if (user != null && user.GetAttributeValue<bool>("isdisabled") == false)
                {
                    bReturnValue = true;
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return bReturnValue;
        }

        public Entity GetMatter(Guid matterId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity matter = null;

            try
            {
                matter = objService.Retrieve("cpdc_matter", matterId, new ColumnSet(new string[] { "cpdc_matterid", "cpdc_billingtimekeeperid" }));
            }
            catch (Exception ex)
            {
                objTracer.Trace("WorkflowUtilities.Getmatter: " + ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return matter;
        }

        public Entity GetOpportunity(Guid oppId, IOrganizationService objService, ITracingService objTracer)
        {
            Entity opp = null;

            try
            {
                opp = objService.Retrieve("opportunity", oppId, new ColumnSet(new string[] { "cxp_regionalmanager", "rg_servicingofficeid", "cxp_industrypracticearea", "ownerid", "cxp_servicerequested", "estimatedvalue", "parentaccountid" }));
            }
            catch (Exception ex)
            {
                objTracer.Trace("WorkflowUtilities.GetOpportunity: " + ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return opp;
        }

        public EntityReference GetOfficeForBillingTimekeeper(EntityReference btkRef, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference officeRef = null;

            try
            {
                Entity btk = objService.Retrieve("systemuser", btkRef.Id, new ColumnSet(new string[] { "systemuserid", "cxp_gloffice" }));

                if (btk != null)
                {
                    officeRef = btk.GetAttributeValue<EntityReference>("cxp_gloffice");
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace("WorkflowUtilities.GetOfficeForBillingTimekeeper: " + ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return officeRef;
        }

        public EntityReference GetOMPForOffice(EntityReference officeRef, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference ompRef = null;

            try
            {
                Entity office = objService.Retrieve("rg_office", officeRef.Id, new ColumnSet(new string[] { "rg_officeid", "cxp_officemanagingpartner" }));

                if (office != null)
                {
                    ompRef = office.GetAttributeValue<EntityReference>("cxp_officemanagingpartner");
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace("WorkflowUtilities.GetOMPForOffice: " + ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return ompRef;
        }

        public List<Entity> GetMattersForAccount(Guid accountId, ColumnSet columns, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> matters = new List<Entity>();

            try
            {
                QueryByAttribute query = new QueryByAttribute("cpdc_matter");
                query.ColumnSet = columns;
                query.AddAttributeValue("cpdc_clientid", accountId);

                EntityCollection results = objService.RetrieveMultiple(query);

                matters = results.Entities.ToList();
            }
            catch (Exception ex)
            {
                objTracer.Trace("WorkflowUtilities.GetMattersForAccount: " + ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return matters;
        }

        public string GetOptionSetTextForValue(OptionSetValue optionSetValue, string entityType, string attributeName, IOrganizationService objService, ITracingService objTracer)
        {
            string label = "";

            if (optionSetValue != null)
            {
                RetrieveAttributeRequest retrieveAttributeRequest = new
                RetrieveAttributeRequest
                {
                    EntityLogicalName = entityType,
                    LogicalName = attributeName
                    // RetrieveAsIfPublished = true
                };

                try
                {
                    // Execute the request.
                    RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)objService.Execute(retrieveAttributeRequest);
                    // Access the retrieved attribute.
                    Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)
                    retrieveAttributeResponse.AttributeMetadata;// Get the current options list for the retrieved attribute.
                    OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();

                    foreach (OptionMetadata oMD in optionList)
                    {
                        if (oMD.Value == optionSetValue.Value)
                        {
                            label = oMD.Label.LocalizedLabels[0].Label.ToString();
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    objTracer.Trace("wfUtilities.GetOptionSetTextForValue: error - " + ex.Message + "\n" + ex.StackTrace);
                }
            }

            return label;
        }

        public void DeleteCXPFollow(Entity cxpFollow, IOrganizationService objService, ITracingService objTracer)
        {
            objService.Delete("cxp_cxpfollow", cxpFollow.Id);
        }

        public string GetUserFirstNameLastName(EntityReference userRef, IOrganizationService objService, ITracingService objTracer)
        {
            string name = "";

            try
            {
                Entity user = objService.Retrieve("systemuser", userRef.Id, new ColumnSet(new string[] { "firstname", "lastname" }));

                if (user != null)
                {
                    name = user.GetAttributeValue<string>("firstname") + " " + user.GetAttributeValue<string>("lastname");
                }
            }
            catch (Exception ex)
            {
                //file.WriteLine("Error: " + ex.Message + "\n" + ex.StackTrace);
                Console.WriteLine("Error: " + ex.Message + "\n" + ex.StackTrace);
            }

            return name;
        }
    }
}
ecutionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            Guid leadId = context.PrimaryEntityId;

            Entity lead = service.Retrieve("lead", leadId, new Microsoft.Xrm.Sdk.Query.ColumnSet(new string[] { "address1_line3" }));

            string userName = lead.GetAttributeValue<string>("address1_line3");

            if (!userName.Equals("Bryan Holmstrom"))
            {
                Guid userId = GetUserId(tracingService, service, userName);
                if (userId != Guid.Empty)
                {
                    AssignLead(tracingService, service, new EntityReference("systemuser", userId), leadId);
                }
            }


            throw new NotImplementedException();
        }

        protected Guid GetUserId(ITracingService tracer, IOrganizationService service, string userName)
        {
            Guid userId = Guid.Empty;

            try
            {
                QueryByAttribute query = new QueryByAttribute("systemuser");
                query.ColumnSet = new ColumnSet(new string[] { "systemuserid" });
                query.AddAttributeValue("fullname", userName);
                EntityCollection results = service.RetrieveMultiple(query);

                userId = results.Entities[0].GetAttributeValue<Guid>("systemuserid");
            }
            catch (Exception e)
            {
                tracer.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw e;
            }

            return userId;
        }

        protected void AssignLead(ITracingService tracer, IOrganizationService service, EntityReference user, Guid leadId)
        {
            try
            {
                Entity lead = new Entity("lead", leadId);
                lead["ownerid"] = user;
                service.Update(lead);
            }
            catch (Exception e)
            {
                tracer.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw e;
            }
        }
    }
}
