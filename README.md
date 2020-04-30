using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;

namespace CXPWorkflows
{
    public class GenerateNewMatterNumber : CodeActivity
    {
        //private IOrganizationService service;
        //private ITracingService tracer;

        [RequiredArgument]
        [Input("Client")]
        [ReferenceTarget("account")]
        public InArgument<EntityReference> Client { get; set; }
        

                                   [RequiredArgument]
                                    [Input("OpportunityProduct")]
                                    [ReferenceTarget("product")]
                                    public InArgument<EntityReference> OpportunityProduct { get; set; }

        [RequiredArgument]
        [Input("MatterYear")]
        public InArgument<string> MatterYear { get; set; }

        [Output("MatterNumber")]
        public OutArgument<string> MatterNumber { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            // Get the context service.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            //service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);

            ITracingService objTracer = executionContext.GetExtension<ITracingService>();
            //tracer = executionContext.GetExtension<ITracingService>();

            EntityReference accountRef = this.Client.Get(executionContext);
            EntityReference serviceRef = this.OpportunityProduct.Get(executionContext);
            string matterYear = this.MatterYear.Get(executionContext);
         
            if (serviceRef == null)
            {
                throw new Exception("Opportunity Product needs to be updated before generating the Matter number.");
            }
            if (matterYear == null)
            {
                throw new Exception("Matter Year needs to be updated before generating the Matter number.");
            }

                // get service area reference, appropriate year and client number
                Entity account = null;
            string accountNumber = String.Empty;

            if (accountRef != null)
            {
                account = GetAccount(accountRef, objService, objTracer);
            }

            if (account != null)
            {
                accountNumber = account.GetAttributeValue<string>("accountnumber");
                objTracer.Trace("accountNumber: " + accountNumber);
            }

            Entity product = null;

            if (serviceRef != null)
            {
                product = GetServiceArea(serviceRef, objService, objTracer);
            }

            //int year = DateTime.Now.Year;
            //if (serviceArea != null && serviceArea.GetAttributeValue<string>("cpdc_servicetype").ToLower().Equals("tax compliance"))
            //{
            //    year = year - 1;
            //}

            //string yearStr = year.ToString().Substring(2, 2);

            string matterCode = product.GetAttributeValue<string>("cxp_mattercode");
            objTracer.Trace("matterCode: " + matterCode);


            string matterNumber = accountNumber + "-" + matterCode + "-" + matterYear;
            objTracer.Trace("matterNumber: " + matterNumber);
          
            bool multipleServicesAllowed = product.GetAttributeValue<bool>("cxp_multipleservices");
            objTracer.Trace("multipleServicesAllowed: " + multipleServicesAllowed);

            while (MatterNumberExists(matterNumber, objService, objTracer))
            {
                objTracer.Trace("matterNumber exists already");
                if (multipleServicesAllowed)
                {
                    int countOfServices = 0;
                    countOfServices = RetrieveCountOfTheService(accountRef.Id, serviceRef.Id, matterYear, objService, objTracer);

                    if (countOfServices < 5)
                    {
                        int matterCodeInt = 0;
                        bool isNumeric = int.TryParse(matterCode, out matterCodeInt);
                        if (isNumeric)
                        {
                            objTracer.Trace("isNumeric = true");
                            matterCodeInt++;
                            matterCode = matterCodeInt.ToString();
                            objTracer.Trace("matterCode: " + matterCode);
                            matterNumber = accountNumber + "-" + matterCode + "-" + matterYear;
                            objTracer.Trace("matterNumber: " + matterNumber);
                        }
                        else
                        {
                            objTracer.Trace("isNumeric = false");
                            string lastCharacter = matterCode.Substring(matterCode.Length - 1, 1);
                            objTracer.Trace("lastCharacter: " + lastCharacter);
                            int lastCharInt = 0;
                            bool isLastNumeric = Int32.TryParse(lastCharacter, out lastCharInt);

                            if (isLastNumeric)
                            {
                                objTracer.Trace("isLastNumeric = true");
                                lastCharInt++;
                            }
                            else
                            {
                                objTracer.Trace("isLastNumeric = false");
                                lastCharInt = 1;
                            }


                            matterCode = matterCode.Substring(0, matterCode.Length - 1) + lastCharInt.ToString();
                            objTracer.Trace("matterCode: " + matterCode);
                            matterNumber = accountNumber + "-" + matterCode + "-" + matterYear;
                            objTracer.Trace("matterNumber: " + matterNumber);
                        }
                    }
                    else
                    {
                        throw new Exception("More than 5 services are not allowed in the same year for an account with the same service requested.");
                    }
                }
                else
                {
                    throw new Exception("Multiple services are not allowed in the same year for an account for the service requested ");
                }

            }

            MatterNumber.Set(executionContext, matterNumber);

        }

        protected int RetrieveCountOfTheService(Guid accountId, Guid productId, string matterYear, IOrganizationService objService, ITracingService objTracer)
        {
            int count = 0;
            QueryByAttribute query = new QueryByAttribute("cpdc_matter");
            query.ColumnSet = new ColumnSet(new string[] { "cpdc_matternumber" });
            query.AddAttributeValue("cpdc_clientid", accountId);
            query.AddAttributeValue("cxp_product", productId);
            query.AddAttributeValue("cxp_matteryeartext", matterYear);

            EntityCollection results = objService.RetrieveMultiple(query);
            if (results != null && results.Entities.Count > 0)
            {
                foreach (Entity matter in results.Entities)
                {
                    if (matter.Attributes.Contains("cpdc_matternumber") && matter.Attributes["cpdc_matternumber"] != null)
                    {
                        count++;
                    }
                }
            }
            return count;
        }

        protected bool MatterNumberExists(string matterNumber, IOrganizationService objService, ITracingService objTracer)
        {
            bool bMatterExists = false;

            try
            {
                QueryByAttribute query = new QueryByAttribute("cpdc_matter");
                query.ColumnSet = new ColumnSet(new string[] { "cpdc_matterid" });
                query.AddAttributeValue("cpdc_matternumber", matterNumber);
                query.AddAttributeValue("statecode", 0);

                EntityCollection results = objService.RetrieveMultiple(query);
                if (results != null && results.Entities.Count > 0)
                {
                    bMatterExists = true;
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return bMatterExists;
        }

        protected Entity GetServiceArea(EntityReference serviceRef, IOrganizationService objService, ITracingService objTracer)
        {
            Entity product = null;

            try
            {
                product = objService.Retrieve("product", serviceRef.Id, new ColumnSet(new string[] { "cxp_mattercode", "productid", "cxp_servicetype", "cxp_multipleservices" }));
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return product;
        }

        protected Entity GetAccount(EntityReference accountRef, IOrganizationService objService, ITracingService objTracer)
        {
            Entity account = null;

            try
            {
                account = objService.Retrieve("account", accountRef.Id, new ColumnSet(new string[] { "accountnumber", "accountid" }));
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                //throw ex;
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }

            return account;
        }
    }
}

