using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Net;
using System.Net.Http;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;

namespace CXPWorkflows
{
    public class PushAccountToIntegrationWorkflow : CodeActivity
    {
		[RequiredArgument]
		[Input("account")]
		[ReferenceTarget("account")]
		public InArgument<EntityReference> account { get; set; }	
        
        protected override void Execute(CodeActivityContext executionContext)
        {
				
		   // Get the context service.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            
            ITracingService objTracer = executionContext.GetExtension<ITracingService>();
            
            WorkflowUtilities wfUtilities = new WorkflowUtilities();
			
			Guid accId = account.Get(executionContext).Id;
           
            if (accId != null)
            {           
                //HttpWebResponse response = null;
                try
                {
                     Entity account = objService.Retrieve("account", accId, new ColumnSet(new string[] { "accountnumber", "cpdc_companyrelationship" }));
                    //Entity account = postAccount;//Added by me MSFT, Instead of using Service.retrieve we are using Post image to get the values.
                    if (account != null && account.GetAttributeValue<OptionSetValue>("cpdc_companyrelationship") != null &&
                        (account.GetAttributeValue<OptionSetValue>("cpdc_companyrelationship").Value == 100000000 ||
                        account.GetAttributeValue<OptionSetValue>("cpdc_companyrelationship").Value == 100000001) &&
                        account.GetAttributeValue<string>("accountnumber") != null &&
                        account.GetAttributeValue<string>("accountnumber") != string.Empty)
                    {           
                        string clientNumber = account.GetAttributeValue<string>("accountnumber");
                        if (clientNumber == null || clientNumber == string.Empty)
                        {
                            throw new Exception("This account does not contain a client number.  Accounts without a client number cannot be integrated with Elite.");
                        }

                        String integration_url = wfUtilities.GetSystemParameter("Account Integration URL", objService, objTracer);
                     
                        integration_url = integration_url + account.GetAttributeValue<string>("accountnumber");
                       
                        WebRequest request = WebRequest.Create(integration_url);
                    
                        request.Timeout = 300000;
                     
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        StreamReader responseReader = new StreamReader(response.GetResponseStream()); // we read response
                      
                        if (responseReader.ReadToEnd().Contains("Success")) //success
                        {
                            objTracer.Trace("Response from Elite Integration Endpoint (Client): " + "SUCCESS");
                        }
                        else //fail, length is 6
                        {
                            objTracer.Trace("Response from Elite Integration Endpoint (Client): " + "FAIL");
                        }

                    }
                } //end of try

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in MyPlug-in.", ex);
                }

                catch (Exception ex)
                {
                    objTracer.Trace("MyPlugin: {0}", ex.ToString());
                    //throw;
                    throw new InvalidPluginExecutionException("MyPlugin: {0}" + ex.ToString());
                }

            }
        }
    }
}
