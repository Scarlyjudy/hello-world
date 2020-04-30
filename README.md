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
    //Description: This custom workflow is triggered by the ui workflow "TriggerEmailAlertBasedOnEstFees" when Est fees >= 25000, on a child opportunity
    //************ The below code retrieves the users : OMP, RMP,Industry Lead, Industry Group Lead, and the users to be cc'd, and passes it to the ui workflow to send the email 
    public class TriggerEmailAlertsBasedOnEstFees : CodeActivity
    {
        //private IOrganizationService service;
        //private ITracingService tracer;
        //private WorkflowUtilities wfUtilities;

        //[Input("Opportunity")]
        //[ReferenceTarget("account")]
        //public InArgument<EntityReference> Account { get; set; }

        //[Input("Contact")]
        //[ReferenceTarget("contact")]
        //public InArgument<EntityReference> Contact { get; set; }

        [RequiredArgument]
        [Input("Opportunity")]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }

        [Output("OMP")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> OMPReference { get; set; }

        [Output("RMP")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> RMPReference { get; set; }

        [Output("IndLead")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> IndLeadReference { get; set; }

        [Output("IndGrpLead")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> IndGrpLeadReference { get; set; }

        [Output("Advisory1")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Adv1Reference { get; set; }

        [Output("Advisory2")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Adv2Reference { get; set; }

        [Output("Advisory3")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Adv3Reference { get; set; }

        [Output("Tax1")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Tax1Reference { get; set; }

        [Output("Tax2")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Tax2Reference { get; set; }

        [Output("Tax3")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Tax3Reference { get; set; }

        [Output("Assurance1")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Asrnc1Reference { get; set; }

        //Independence Notification Users

        [Output("IndependenceNotificationUsers")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> IndptAppReference { get; set; }

        //conflict Notification Users

        [Output("ConflictNotificationUsers")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> ConflictReference { get; set; }


        //Cannabis Assurance Notification Users

        [Output("CannabisAssuranceUsers")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CannabisAssurReference { get; set; }

        //
        //Cannabis tax Director  Notification Users

        [Output("CannabisTaxDir")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CTaxDirReference { get; set; }
        //Cannabis tax Partner  Notification Users

        [Output("CannabisTaxPartner")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CTaxPartnerReference { get; set; }

        //Financial Services  Notification Users

        [Output("FSIusers")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> FSIReference { get; set; }

        //Securities  Notification Users

        [Output("SECusers")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> SECReference { get; set; }

        //

        [Output("Assurance2")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Asrnc2Reference { get; set; }

        [Output("Assurance3")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> Asrnc3Reference { get; set; }

        [Output("CC1")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC1Reference { get; set; }

        [Output("CC2")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC2Reference { get; set; }

        [Output("CC3")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC3Reference { get; set; }

        [Output("CC4")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC4Reference { get; set; }

        [Output("CC5")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC5Reference { get; set; }

        [Output("CC6")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC6Reference { get; set; }

        [Output("CC7")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC7Reference { get; set; }

        [Output("CC8")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC8Reference { get; set; }

        [Output("CC9")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC9Reference { get; set; }

        [Output("CC10")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CC10Reference { get; set; }

        [Output("CCPortfolioManager")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> PFMgrReference { get; set; }

        [Output("CREManager1")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CREMgr1Reference { get; set; }
        [Output("CREManager2")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> CREMgr2Reference { get; set; }

        [Output("MarketingLead")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> MktingLeadReference { get; set; }

        bool physicalOffice = true;
        bool advisoryVirtualOffice = false;
        bool taxVirtualOffice = false;
        bool assuranceVirtualOffice = false;
        EntityReference PFMgrRef = null;

        protected override void Execute(CodeActivityContext executionContext)
        {
            // Get the context service.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.InitiatingUserId);

            ITracingService objTracer = executionContext.GetExtension<ITracingService>();

            WorkflowUtilities wfUtilities = new WorkflowUtilities();

            //Step 1: Retrieve the opportunity
            // EntityReference oppRef = this.Opportunity.Get(executionContext);
            Guid oppId = Opportunity.Get(executionContext).Id;
            Entity opp = null;

            if (oppId != null)
            {
                // opp = GetOpp(oppRef, objService, objTracer);
                opp = wfUtilities.GetOpportunity(oppId, objService, objTracer);

                objTracer.Trace("opp retrieved:" + oppId);

                //Step 2: Retrieve the users - to send email
                EntityReference RMP = null;
                EntityReference OMP = null;
                EntityReference IndustryLead = null;
                EntityReference IndustryGroupLead = null;
                EntityReference Advisory1 = null;
                EntityReference Advisory2 = null;
                EntityReference Advisory3 = null;
                EntityReference Tax1 = null;
                EntityReference Tax2 = null;
                EntityReference Tax3 = null;
                EntityReference Assurance1 = null;
                EntityReference IndependenceNotificationUsers = null;
                EntityReference ConflictNotificationUsers = null;
                EntityReference CannabisAssuranceUsers = null;

                EntityReference CannabisTaxDir = null;
                EntityReference CannabisTaxPartner = null;
                EntityReference FSIusers = null;
                EntityReference SECusers = null;

                EntityReference Assurance2 = null;
                EntityReference Assurance3 = null;
                EntityReference cc1 = null;
                EntityReference cc2 = null;
                EntityReference cc3 = null;
                EntityReference cc4 = null;
                EntityReference cc5 = null;
                EntityReference cc6 = null;
                EntityReference cc7 = null;
                EntityReference cc8 = null;
                EntityReference cc9 = null;
                EntityReference cc10 = null;
                EntityReference CREMgr1 = null;
                EntityReference CREMgr2 = null;
                EntityReference MktingLead = null;

                if (opp != null)
                {
                    //retrieve RMP OMP Ind Lead, Ind Grp Lead
                    objTracer.Trace("opp not null, retrieving users" + opp);
                    RMP = retrieveRMP(opp, objService, objTracer);
                    objTracer.Trace("RMP retrieved");
                    OMP = retrieveOMP(opp, objService, objTracer);
                    objTracer.Trace("OMP retrieved");
                    IndustryLead = retrieveIndLead(opp, objService, objTracer);
                    objTracer.Trace("IndustryLead retrieved");
                    IndustryGroupLead = retrieveIndGrpLead(opp, objService, objTracer);
                    objTracer.Trace("IndustryGroupLead retrieved");

                    //retrieve users with Major opp alert flag
                    List<Entity> users = retrieveUsersToBeCopiedInTheEmail(opp, objService, objTracer);
                    objTracer.Trace("cc users retrieved");
                    if (users != null && users.Count > 0)
                    {
                        objTracer.Trace("cc users not null, retrieving cc users");
                        cc1 = users[0].ToEntityReference();
                        if (users.Count > 1)
                        {
                            cc2 = users[1].ToEntityReference();
                        }
                        if (users.Count > 2)
                        {
                            cc3 = users[2].ToEntityReference();
                        }
                        if (users.Count > 3)
                        {
                            cc4 = users[3].ToEntityReference();
                        }
                        if (users.Count > 4)
                        {
                            cc5 = users[4].ToEntityReference();
                        }
                        if (users.Count > 5)
                        {
                            cc6 = users[5].ToEntityReference();
                        }
                    }

                    //retrieve users with Major opp alert flag for est fee >= 100K
                    if (opp.GetAttributeValue<Money>("estimatedvalue").Value >= 100000) { 
                        List<Entity> users100K = retrieveUsersToBeCopiedInTheEmailForGEQ100K(opp, objService, objTracer);
                        objTracer.Trace("cc users retrieved for >= 100K");
                        if (users100K != null && users100K.Count > 0)
                        {
                            objTracer.Trace("cc users not null, retrieving cc users -  users100K");
                            cc7 = users100K[0].ToEntityReference();
                            if (users100K.Count > 1)
                            {
                                cc8 = users100K[1].ToEntityReference();
                            }
                            if (users100K.Count > 2)
                            {
                                cc9 = users100K[2].ToEntityReference();
                            }
                            if (users100K.Count > 3)
                            {
                                cc10 = users100K[3].ToEntityReference();
                            }

                        }
                    }
                    //retrieve CRE manager if Pursuit Lead or Lead Gen 1 is a business development (CRE) team member
                    if ((PursuitLeadIsCRETeamMember(opp, objService, objTracer) ) || ( leadGenIsCRETeamMember(opp, objService, objTracer)))
                    {
                        objTracer.Trace("Retrieving CRE managers as Pursuit Lead and or Lead generator 1 is a CRE team member");
                        List<Entity> CRE = retrieveCREManagers(opp, objService, objTracer);
                        objTracer.Trace("CRE Managers retrieved");
                       
                        if (CRE != null && CRE.Count > 0)
                        {
                            objTracer.Trace("CRE managers not null, retrieving CRE managers");
                            objTracer.Trace("CRE.Count:" + CRE.Count);
                            CREMgr1 = CRE[0].ToEntityReference();
                            if (CRE.Count > 1)
                            {
                                CREMgr2 = CRE[1].ToEntityReference();
                            }
                        }
                    }

                    //retrieve Marketing lead from PRODUCT //update
                    MktingLead = retrieveMktingLead(opp, objService, objTracer);

                    //retrieve Advisory alert users if virtual office and advisory region
                    if (advisoryVirtualOffice == true)
                    {
                        List<Entity> adv = retrieveAdvisoryAlertUsers(opp, objService, objTracer);
                        objTracer.Trace("Advisory alert users retrieved");
                        if (adv != null && adv.Count > 0)
                        {
                            objTracer.Trace("Advisory users not null, retrieving advisory users");
                            Advisory1 = adv[0].ToEntityReference();
                            if (adv.Count > 1)
                            {
                                Advisory2 = adv[1].ToEntityReference();
                            }
                            if (adv.Count > 2)
                            {
                                Advisory3 = adv[2].ToEntityReference();
                            }
                        }
                    }

                    //retrieve Tax alert users if virtual office and Tax region
                    if (taxVirtualOffice == true)
                    {
                        List<Entity> tax = retrieveTaxAlertUsers(opp, objService, objTracer);
                        objTracer.Trace("Tax alert users retrieved");
                        if (tax != null && tax.Count > 0)
                        {
                            objTracer.Trace("Advisory users not null, retrieving advisory users");
                            Tax1 = tax[0].ToEntityReference();
                            if (tax.Count > 1)
                            {
                                Tax2 = tax[1].ToEntityReference();
                            }
                            if (tax.Count > 2)
                            {
                                Tax3 = tax[2].ToEntityReference();
                            }
                        }
                    }

                    //retrieve Assurance alert users if virtual office and Tax region
                    if (assuranceVirtualOffice == true)
                    {
                        List<Entity> asrnc = retrieveAssuranceAlertUsers(opp, objService, objTracer);
                        objTracer.Trace("Assurance alert users retrieved");
                        if (asrnc != null && asrnc.Count > 0)
                        {
                            objTracer.Trace("Advisory users not null, retrieving advisory users");
                            Assurance1 = asrnc[0].ToEntityReference();
                            if (asrnc.Count > 1)
                            {
                                Assurance2 = asrnc[1].ToEntityReference();
                            }
                            if (asrnc.Count > 2)
                            {
                                Assurance3 = asrnc[2].ToEntityReference();
                            }
                        }
                    }
                    //


                    //retrieve Independence alert users  
                   // if (assuranceVirtualOffice == true)
                    //{
                        List<Entity> IndptAppref = retrieveIndepetAlertUsers(opp, objService, objTracer);
                        objTracer.Trace("INDpT alert users retrieved");
                        if (IndptAppref != null && IndptAppref.Count > 0)
                        {
                            objTracer.Trace("INDPT users not null, retrieving INDPT users");
                           IndependenceNotificationUsers = IndptAppref[0].ToEntityReference();
                         //if (IndptAppref.Count > 1)
                         //{
                         //    IndependenceNotificationUsers1 = IndptAppref[1].ToEntityReference();
                         //}
                         //if (IndptAppref.Count > 2)
                         //{
                         //    IndependenceNotificationUsers2 = IndptAppref[2].ToEntityReference();
                         //}
                    
                        }


                    //retrieve conflict alert users  
                    // if (assuranceVirtualOffice == true)
                    //{
                    List<Entity> conflictusers = retrieveConflictAlertUsers(opp, objService, objTracer);
                    objTracer.Trace("Conflict alert users retrieved");
                    if (conflictusers != null && conflictusers.Count > 0)
                    {
                        objTracer.Trace("Conflict alert users not null, retrieving Conflict users");
                        ConflictNotificationUsers = conflictusers[0].ToEntityReference();
                        //if (conflictusers.Count > 1)
                        //{
                        //    ConflictNotificationUsers = conflictUsers[1].ToEntityReference();
                        //}
                        //if (conflictusers.Count > 2)
                        //{
                        //    ConflictNotificationUsers = conflictUsers[2].ToEntityReference();
                        //}
                    }
                    //}

                    //retrieve Cannabis-Assurance alert users  
                    // if (assuranceVirtualOffice == true)
                    //{
                    List<Entity> canAssurance = retrieveCannabisAssuranceAlertUsers(opp, objService, objTracer);
                    objTracer.Trace("Cannabis-Assurance alert users retrieved");
                    if (canAssurance!= null && canAssurance.Count > 0)
                    {
                        objTracer.Trace("Cannabis Assurance  users not null, retrieving Cannabis-Assurance  users");
                        CannabisAssuranceUsers = canAssurance[0].ToEntityReference();
                        //if (canAssurance.Count > 1)
                        //{
                        //    CannabisAssuranceUsers = canAssurance[1].ToEntityReference();
                        //}
                        //if (canAssurance.Count > 2)
                        //{
                        //    CannabisAssuranceUsers = canAssurance[2].ToEntityReference();
                        //}
                    }
                    //}


                    //retrieve Cannabis-TAX Dir alert users  
                    // if (assuranceVirtualOffice == true)
                    //{
                    List<Entity> CtaxDir = retrieveCannabisTaxDirAlertUsers(opp, objService, objTracer);
                    objTracer.Trace("Cannabis Tax Dir alert users retrieved");
                    if (CtaxDir != null && CtaxDir.Count > 0)
                    {
                        objTracer.Trace("Cannabis Tax Dir users not null, retrieving Cannabis Tax Dir  users");
                        CannabisTaxDir = CtaxDir[0].ToEntityReference();
                        //if (CtaxDir.Count > 1)
                        //{
                        //    CannabisTaxDir = CtaxDir[1].ToEntityReference();
                        //}
                        //if (IndptAppref.Count > 2)
                        //{
                        //    CannabisTaxDir = CtaxDir[2].ToEntityReference();
                        //}
                    }
                    //}

                    //retrieve Cannabis TAX Partner alert users  
                    // if (assuranceVirtualOffice == true)
                    //{
                    List<Entity> CtaxPartner = retrieveCannabisTaxPartAlertUsers(opp, objService, objTracer);
                    objTracer.Trace("Cannabis taxPartner alert users retrieved");
                    if (CtaxPartner != null && CtaxPartner.Count > 0)
                    {
                        objTracer.Trace("Cannabis taxPartner  users not null, retrieving Cannabis taxPartner users");
                        CannabisTaxPartner = CtaxPartner[0].ToEntityReference();
                        //if (CtaxPartner.Count > 1)
                        //{
                        //    CannabisTaxPartner = CtaxPartner[1].ToEntityReference();
                        //}
                        //if (CtaxPartner.Count > 2)
                        //{
                        //    CannabisTaxPartner = CtaxPartner[2].ToEntityReference();
                        //}
                    }
                    //}

                    //retrieve FSI alert users  
                    // if (assuranceVirtualOffice == true)
                    //{
                    List<Entity>  FsiAlertUsers = retrieveFSIAlertUsers(opp, objService, objTracer);
                    objTracer.Trace("Fsi alert users retrieved");
                    if (FsiAlertUsers != null && FsiAlertUsers.Count > 0)
                    {
                        objTracer.Trace("FSI  users not null, retrieving FSI  users");
                        FSIusers = FsiAlertUsers[0].ToEntityReference();
                        //if (FsiAlertUsers.Count > 1)
                        //{
                        //    FSIusers = FsiAlertUsers[1].ToEntityReference();
                        //}
                        //if (FsiAlertUsers.Count > 2)
                        //{
                        //    FSIusers = FsiAlertUsers[2].ToEntityReference();
                        //}
                    }
                    //}

                    //retrieve SEC alert users  
                    // if (assuranceVirtualOffice == true)
                    //{
                    List<Entity> SecAlertUsers = retrieveCannabisAssuranceAlertUsers(opp, objService, objTracer);
                    objTracer.Trace("SEC alert users retrieved");
                    if (SecAlertUsers != null && SecAlertUsers.Count > 0)
                    {
                        objTracer.Trace("SEC  users not null, retrieving SEC users");
                        SECusers = SecAlertUsers[0].ToEntityReference();
                        //if (SecAlertUsers.Count > 1)
                        //{
                        //    SECusers = SecAlertUsers[1].ToEntityReference();
                        //}
                        //if (SecAlertUsers.Count > 2)
                        //{
                        //    SecUsers = SecAlertUsers[2].ToEntityReference();
                        //}
                    }
                    //}

                    //set output variables:
                    RMPReference.Set(executionContext, RMP);
                    OMPReference.Set(executionContext, OMP);
                    IndLeadReference.Set(executionContext, IndustryLead);
                    IndGrpLeadReference.Set(executionContext, IndustryGroupLead);
                    Adv1Reference.Set(executionContext, Advisory1);
                    Adv2Reference.Set(executionContext, Advisory2);
                    Adv3Reference.Set(executionContext, Advisory3);
                    Tax1Reference.Set(executionContext, Tax1);
                    Tax2Reference.Set(executionContext, Tax2);
                    Tax3Reference.Set(executionContext, Tax3);
                    Asrnc1Reference.Set(executionContext, Assurance1);

                    IndptAppReference.Set(executionContext, IndependenceNotificationUsers);
                    ConflictReference.Set(executionContext, ConflictNotificationUsers);
                    CannabisAssurReference.Set(executionContext, CannabisAssuranceUsers);

                    CTaxDirReference.Set(executionContext, CannabisTaxDir);
                    CTaxPartnerReference.Set(executionContext, CannabisTaxPartner);
                    FSIReference.Set(executionContext, FSIusers);
                    SECReference.Set(executionContext, SECusers);

                    Asrnc2Reference.Set(executionContext, Assurance2);
                    Asrnc3Reference.Set(executionContext, Assurance3);
                    CC1Reference.Set(executionContext, cc1);
                    CC2Reference.Set(executionContext, cc2);
                    CC3Reference.Set(executionContext, cc3);
                    CC4Reference.Set(executionContext, cc4);
                    CC5Reference.Set(executionContext, cc5);
                    CC6Reference.Set(executionContext, cc6);
                    CC7Reference.Set(executionContext, cc7);
                    CC8Reference.Set(executionContext, cc8);
                    CC9Reference.Set(executionContext, cc9);
                    CC10Reference.Set(executionContext, cc10);
                    PFMgrReference.Set(executionContext, PFMgrRef);
                    CREMgr1Reference.Set(executionContext, CREMgr1);
                    CREMgr2Reference.Set(executionContext, CREMgr2);
                    MktingLeadReference.Set(executionContext, MktingLead);

                    //Step 3 : Send email

                    //Email is sent through the ui workflow
                }
            }

        }

        //retrieve RMP 
        protected EntityReference retrieveRMP(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference RMPRef = null;
            try
            {
                if (opp != null)
                {
                    RMPRef = opp.GetAttributeValue<EntityReference>("cxp_regionalmanager");
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return RMPRef;
        }

        //retrieve OMP 
        protected EntityReference retrieveOMP(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference OMPRef = null;
            EntityReference SvcingOfcRef = null;
            Entity SvcingOfc = null;
            try
            {
                if (opp != null)
                {
                    SvcingOfcRef = opp.GetAttributeValue<EntityReference>("rg_servicingofficeid");
                }
                if (SvcingOfcRef != null)
                {
                    SvcingOfc = objService.Retrieve("rg_office", SvcingOfcRef.Id, new ColumnSet(new string[] { "cxp_officemanagingpartner", "cxp_physicaloffice", "cxp_region" }));
                    OMPRef = SvcingOfc.GetAttributeValue<EntityReference>("cxp_officemanagingpartner");

                    physicalOffice = SvcingOfc.GetAttributeValue<bool>("cxp_physicaloffice");
                    if (physicalOffice == false)  //if virtual offcice
                    {
                        if (SvcingOfc.GetAttributeValue<EntityReference>("cxp_region").Name == "Advisory")
                        {
                            advisoryVirtualOffice = true;
                        }
                        if (SvcingOfc.GetAttributeValue<EntityReference>("cxp_region").Name == "National Tax")
                        {
                            taxVirtualOffice = true;
                        }
                        //if (SvcingOfc.GetAttributeValue<EntityReference>("cxp_region").Name == "") 
                        //{
                        //    assuranceVirtualOffice = true;
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return OMPRef;
        }

        //retrieve Ind Lead
        protected EntityReference retrieveIndLead(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference IndLeadRef = null;
            EntityReference IndustryRef = null;
            Entity Industry = null;
            EntityReference AccRef = null;
            Entity Acc = null;
            try
            {
                if (opp != null)
                {
                    //IndustryRef = opp.GetAttributeValue<EntityReference>("cxp_industrypracticearea");
                    AccRef = opp.GetAttributeValue<EntityReference>("parentaccountid");
                    Acc = objService.Retrieve("account", AccRef.Id, new ColumnSet(new string[] { "cxp_crindustrypracticearea1" }));
                    IndustryRef = Acc.GetAttributeValue<EntityReference>("cxp_crindustrypracticearea1");
                }
                if (IndustryRef != null)
                {
                    Industry = objService.Retrieve("cpdc_industrypracticearea", IndustryRef.Id, new ColumnSet(new string[] { "cxp_industryleader", "cxp_portfoliomanager" }));
                    IndLeadRef = Industry.GetAttributeValue<EntityReference>("cxp_industryleader");
                    PFMgrRef = Industry.GetAttributeValue<EntityReference>("cxp_portfoliomanager");

                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return IndLeadRef;
        }

        //retrieve Ind Grp Lead
        protected EntityReference retrieveIndGrpLead(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference IndGrpLeadRef = null;
            EntityReference IndustryRef = null;
            EntityReference IndGrpRef = null;
            Entity Industry = null;
            Entity IndustryGrp = null;
            EntityReference AccRef = null;
            Entity Acc = null;
            try
            {
                if (opp != null)
                {
                    //IndustryRef = opp.GetAttributeValue<EntityReference>("cxp_industrypracticearea");
                    AccRef = opp.GetAttributeValue<EntityReference>("parentaccountid");
                    Acc = objService.Retrieve("account", AccRef.Id, new ColumnSet(new string[] { "cxp_crindustrypracticearea1" }));
                    IndustryRef = Acc.GetAttributeValue<EntityReference>("cxp_crindustrypracticearea1");
                }
                if (IndustryRef != null)
                {
                    Industry = objService.Retrieve("cpdc_industrypracticearea", IndustryRef.Id, new ColumnSet(new string[] { "cxp_industrygroup" }));
                    IndGrpRef = Industry.GetAttributeValue<EntityReference>("cxp_industrygroup");
                    IndustryGrp = objService.Retrieve("cxp_industrygroup", IndGrpRef.Id, new ColumnSet(new string[] { "cxp_leadindustrypartner" }));
                    IndGrpLeadRef = IndustryGrp.GetAttributeValue<EntityReference>("cxp_leadindustrypartner");
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return IndGrpLeadRef;
        }

        //retrieve Advisory alert users if virtual office and advisory region
        protected List<Entity> retrieveAdvisoryAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> adv = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_advisory_alert", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    adv = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return adv;

        }

        //retrieve Tax alert users if virtual office and Tax region
        protected List<Entity> retrieveTaxAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> tax = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_taxalert", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    tax = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return tax;

        }

        //retrieve Assurance alert users if virtual office and advisory region
        protected List<Entity> retrieveAssuranceAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> asrnc = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_assurancealert", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    asrnc = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return asrnc;

        }


        //TEST

        //retrieveIndepetAlertUsers  
        protected List<Entity> retrieveIndepetAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> IndptAppref = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_nationalindependenceapprover", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    IndptAppref = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return IndptAppref;

        }

        //


        //retrieve Conflict Alert Users alert users if virtual office and advisory region
        protected List<Entity> retrieveConflictAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> conflictusers = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_conflictofinterestnotifications", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    conflictusers = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return conflictusers;

        }


        //Retrieve Cannabis Assurance Alerts users  
        protected List<Entity> retrieveCannabisAssuranceAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> CanAssurance = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_cannabisassuranceapprover", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    CanAssurance = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return CanAssurance;

        }

        //


        //Retrieve Cannabis Tax Partners  Alerts users  
        protected List<Entity> retrieveCannabisTaxPartAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> CtaxPartner = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_cannabistaxpartner", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    CtaxPartner = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return CtaxPartner;

        }

        //

        //Retrieve Cannabis Tax Director Alerts users  
        protected List<Entity> retrieveCannabisTaxDirAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> CtaxDir = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_cannabistaxpracticedirector", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    CtaxDir = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return CtaxDir;

        }

        //

        //Retrieve SEC Alerts users  
        protected List<Entity> retrieveSecAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> SecAlertUsers = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_secassurance", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    SecAlertUsers = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return SecAlertUsers;

        }

        //

        //Retrieve FSI Assurance Alerts users  
        protected List<Entity> retrieveFSIAlertUsers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> FsiAlertUsers = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_financialservices", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    FsiAlertUsers = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return FsiAlertUsers;

        }

        //
        //retrieve users with Major opp alert flag
        protected List<Entity> retrieveUsersToBeCopiedInTheEmail(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> users = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_large_opp_alert", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    users = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return users;
        }

        //retrieve users with Major opp alert flag for est fee >= 100K
        protected List<Entity> retrieveUsersToBeCopiedInTheEmailForGEQ100K(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> users = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_majoroppalert100k", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    users = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return users;
        }

        //retrieve CRE managers if Pursuit Lead is a business development (CRE) team member
        protected List<Entity> retrieveCREManagers(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            List<Entity> cre = new List<Entity>();
            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("cxp_cremanager", ConditionOperator.Equal, true);

                EntityCollection results = objService.RetrieveMultiple(query);

                if (results != null)
                {
                    cre = results.Entities.ToList();
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return cre;
         }

        //verify if pursuit lead is cre team member
        protected bool PursuitLeadIsCRETeamMember(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference PursuitLeadRef = null;
             
            bool IsCRETeamMember = false;
            try
            {
                if (opp != null)
                {
                    PursuitLeadRef = opp.GetAttributeValue<EntityReference>("ownerid");
                    
                    objTracer.Trace("PursuitLead Ref:" + PursuitLeadRef);
                   

                    if (PursuitLeadRef != null)
                    {
                        Entity user = objService.Retrieve("systemuser", PursuitLeadRef.Id, new ColumnSet(new string[] { "firstname", "lastname", "cxp_cre" }));
                        if (user != null)
                        {
                           IsCRETeamMember = user.GetAttributeValue<bool>("cxp_cre");
                            objTracer.Trace("userIsCreTeammem" + IsCRETeamMember);
                        }
                    }

                  
                    
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            objTracer.Trace("userIsCreTeammem" );
            return IsCRETeamMember;
        }

        //verify if   lead Generator 1 is cre team member
        protected bool leadGenIsCRETeamMember(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
           
            EntityReference leadGen1Ref = null;
            bool IsCRETeamMember = false;
            try
            {
                if (opp != null)
                {
                   
                    leadGen1Ref = opp.GetAttributeValue<EntityReference>("cxp_leadgenerator1");

                    
                    objTracer.Trace("leadGen1 Ref:" + leadGen1Ref);

                    if (leadGen1Ref != null)
                    {
                        Entity user = objService.Retrieve("systemuser", leadGen1Ref.Id, new ColumnSet(new string[] { "firstname", "lastname", "cxp_cre" }));
                        if (user != null)
                        {
                            IsCRETeamMember = user.GetAttributeValue<bool>("cxp_cre");
                            objTracer.Trace("userIsCreTeammem" + IsCRETeamMember);
                        }
                    }



                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            objTracer.Trace("userIsCreTeammem");
            return IsCRETeamMember;
        }

        //retrieve Marketing Lead
        protected EntityReference retrieveMktingLeadTEST(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference MktLeadRef = null;
            EntityReference SvcAreaRef = null;
            Entity SvcArea = null;
            int SvcType;
            int assurance = 280410001;
            try
            {
                if (opp != null)
                {
                    SvcAreaRef = opp.GetAttributeValue<EntityReference>("cxp_servicerequested");  
                }
                if (SvcAreaRef != null)
                {
                    SvcArea = objService.Retrieve("rg_servicearea", SvcAreaRef.Id, new ColumnSet(new string[] { "cxp_marketinglead","cxp_servicecategory" }));
                    MktLeadRef = SvcArea.GetAttributeValue<EntityReference>("cxp_marketinglead");
                    SvcType = SvcArea.GetAttributeValue<OptionSetValue>("cxp_servicecategory").Value;
                    if (SvcType == assurance )
                    {
                        assuranceVirtualOffice = true;
                    }
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return MktLeadRef;
        }


        /*
        //retrieve Marketing Lead From Existing Product
        protected EntityReference retrieveMktingLead(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            EntityReference MktLeadRef = null;
            EntityReference opproductRef = null;
            Entity cxpProduct = null;
            int srvceCatg;
            //int assurance = 280410001;
            try
            {
                if (opp != null)
                {
                    opproductRef = opp.GetAttributeValue<EntityReference>("cxp_opportunityproduct"); //
                }
                if (opproductRef != null)
                {
                   // Entity oppProduct = null;
                   // EntityReference productRef = null;
                   // productRef = opp.GetAttributeValue<EntityReference>("productid");

                    Entity Opp_products = objService.Retrieve("cxp_opportunityproduct", opproductRef.Id, new ColumnSet(new string[] { "productid"}));
                    EntityReference productRef = Opp_products.GetAttributeValue<EntityReference>("productid");

                    cxpProduct = objService.Retrieve("product", productRef.Id, new ColumnSet(new string[] { "cxp_marketinglead", "cxp_servicecategory" }));
                    MktLeadRef = cxpProduct.GetAttributeValue<EntityReference>("cxp_marketinglead");
                    srvceCatg = cxpProduct.GetAttributeValue<OptionSetValue>("cxp_servicecategory").Value;

                    if (cxpProduct != null && srvceCatg == 280410001)
                    {
                        assuranceVirtualOffice = true;
                    }
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return MktLeadRef;
        }   */

        //retrieve Marketing Lead From Existing Product
        protected EntityReference retrieveMktingLead(Entity opp, IOrganizationService objService, ITracingService objTracer)
        {
            //throw new InvalidPluginExecutionException("TEXT Ex ----");
            objTracer.Trace("Entering FXN-----Retrieving mktLead");
            EntityReference OpproductRef = null;
            Entity CxpProduct = null;
            EntityReference MktLeadRef = null;
            // int SrvceCatg;
            int assurance = 280410001;
            try
            {
                //throw new InvalidPluginExecutionException("TEXT Ex ----");
                //objTracer.Trace("Try -----");
                if (opp != null)
                {
                    objTracer.Trace("Setting ------- " + opp);
                    OpproductRef = opp.GetAttributeValue<EntityReference>("cxp_opportunityproduct");

                    objTracer.Trace("Opp Product Ref:" + OpproductRef);//
                }
                if (OpproductRef != null)
                {

                    objTracer.Trace("Setting Product --------");
                    Entity Opp_products = objService.Retrieve("cxp_opportunityproduct", OpproductRef.Id, new ColumnSet(new string[] { "productid" }));
                    EntityReference ProductRef = Opp_products.GetAttributeValue<EntityReference>("productid");
                    objTracer.Trace("Setting  --------" + Opp_products);
                    CxpProduct = objService.Retrieve("product", ProductRef.Id, new ColumnSet(new string[] { "cxp_marketinglead", "cxp_servicecategory" }));
                    MktLeadRef = CxpProduct.GetAttributeValue<EntityReference>("cxp_marketinglead");
                    objTracer.Trace("Setting maktLeadRef  --------" + MktLeadRef);
                    int SrvceCatg = CxpProduct.GetAttributeValue<OptionSetValue>("cxp_servicecategory").Value;

                    objTracer.Trace("srv category  check ----" + SrvceCatg);

                    if (CxpProduct != null && SrvceCatg == assurance)
                    {
                        objTracer.Trace("cxp prdct check ----" + CxpProduct);
                        assuranceVirtualOffice = true;
                    }
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "MKT LEAD \n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return MktLeadRef;
        }

    }
}
