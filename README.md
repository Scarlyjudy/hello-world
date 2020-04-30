using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
//using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;

namespace CXPUtilities
{
    public class CXPUtilities
    {
        private IOrganizationService service;
        private ITracingService tracer;

        public IOrganizationService Service
        {
            get { return service; }
            set { service = value; }
        }
        
        public ITracingService Tracer
        {
            get { return tracer; }
            set { tracer = value; }
        }

        public Entity GetAccount(Guid accountId)
        {
            Entity account = null;

            try
            {
                account = service.Retrieve("account", accountId, new ColumnSet(new string[] { "createdby" }));
                tracer.Trace("GetAccount: account = " + account != null ? account.ToString() : "null.");
            }
            catch(Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return account;
        }

        public Entity GetContact(Guid contactId)
        {
            Entity contact = null;

            try
            {
                contact = service.Retrieve("contact", contactId, new ColumnSet(new string[] { "createdby" }));
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return contact;
        }

        public Entity GetUser(Guid userId)
        {
            Entity user = null;

            try
            {
                user = service.Retrieve("systemuser", userId, new ColumnSet(new string[] { "businessunitid" }));
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return user;
        }

        public EntityReference GetBUOwnerRef(string buName)
        {
            EntityReference buOwnerRef = null;

            try
            {
                string systemParameterName = buName + " User";
                QueryByAttribute query = new QueryByAttribute("cxp_systemparameter");
                query.AddAttributeValue("cxp_name", systemParameterName);
                query.ColumnSet = new ColumnSet(new string[] { "cxp_value" });

                EntityCollection results = service.RetrieveMultiple(query);

                if (results != null)
                {
                    string buUserName = results.Entities[0].GetAttributeValue<string>("cxp_value");
                    Entity buUser = GetUserByName(buUserName);
                    if (buUser != null)
                    {
                        buOwnerRef = new EntityReference("systemuser", buUser.GetAttributeValue<Guid>("systemuserid"));
                        buOwnerRef.Name = buUser.GetAttributeValue<string>("fullname");
                    }
                }
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return buOwnerRef;
        }

        public Entity GetUserByName(string name)
        {
            Entity user = null;

            try
            {
                QueryByAttribute query = new QueryByAttribute("systemuser");
                query.AddAttributeValue("fullname", name);
                query.ColumnSet = new ColumnSet(new string[] { "systemuserid", "fullname" });

                EntityCollection results = service.RetrieveMultiple(query);

                if (results != null)
                {
                    user = results.Entities[0];
                }
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return user;
        }

        public void AddUserToRecordAccessTeam(Guid userId, string recordType, Guid recordId)
        {
            try
            {
                AddUserToRecordTeamRequest request = new AddUserToRecordTeamRequest();
                request.SystemUserId = userId;
                request.Record = new EntityReference(recordType, recordId);
                request.TeamTemplateId = GetAccessTeamTemplateIdByName(recordType);

                service.Execute(request);
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }
        }

        public Guid GetAccessTeamTemplateIdByName(String recordType)
        {
            Guid templateId = Guid.Empty;
            string accessTeamTemplateName = GetAccessTeamTemplateName(recordType);

            if (accessTeamTemplateName != String.Empty)
            {
                try
                {
                    QueryExpression teamQuery = new QueryExpression
                    {
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

                    var results = service.RetrieveMultiple(teamQuery).Entities;

                    if (results.Count != 1)
                        throw new Exception("Can't retrieve Team Access Template for " + recordType);
                    else
                    {
                        templateId = results[0].Id;
                    }
                }
                catch (Exception ex)
                {
                    tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                    throw ex;
                }
            }
            
            return templateId;
        }

        public string GetAccessTeamTemplateName(string recordType)
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

                    EntityCollection results = service.RetrieveMultiple(query);

                    if (results != null)
                    {
                        teamName = results.Entities[0].GetAttributeValue<string>("cxp_value");
                    }
                }
                catch (Exception ex)
                {
                    tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                    throw ex;
                }
            }
            
            return teamName;
        }

        public void CreateFollowOfRecordForCreatedByUser(Guid userId, string recordType, Guid recordId)
        {
            try
            {
                Entity postFollow = new Entity("postfollow");
                postFollow.Attributes["regardingobjectid"] = new EntityReference(recordType, recordId);
                postFollow.Attributes["ownerid"] = new EntityReference("systemuser", userId);
                service.Create(postFollow);
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }
        }

        public Entity GetAccountCommunicationOptions(Guid accountId)
        {
            Entity account = null;

            try
            {
                account = service.Retrieve("account", accountId, new ColumnSet(new string[] { "donotbulkemail","donotbulkpostalmail","donotemail","donotphone","donotfax","donotpostalmail","donotsendmm"}));
                tracer.Trace("GetAccount: account = " + account != null ? account.ToString() : "null.");
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return account;
        }

        public List<Entity> GetContactsForAccount(Guid accountId)
        {
            List<Entity> contacts = new List<Entity>();

            try
            {
                QueryByAttribute query = new QueryByAttribute("contact");
                query.AddAttributeValue("parentcustomerid", accountId);
                query.ColumnSet = new ColumnSet(new string[] { "contactid" });

                EntityCollection results = service.RetrieveMultiple(query);

                if (results != null)
                {
                    contacts = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                tracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw ex;
            }

            return contacts;
        }

        public bool boolCheckIfUserIsAlreadyOnAccessTeam(Guid userId, string recordType, Guid recordId)
        {
            bool returnValue = false;

            return returnValue;
        }
    }
}
