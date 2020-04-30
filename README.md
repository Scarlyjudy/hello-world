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
if (typeof (CR) === "undefined") {
    CR = {
        _namespace: true
    };
}
if (typeof (CR.Opportunity) === "undefined") {
    CR.Opportunity = {
        _namespace: true
    };
}
CR.Opportunity.Form = (function () {
    "use strict";
    var self = {};
    //The onLoad method 
    self.onLoad = function (executionContext) {
        try {
            debugger;
            setFormContextFromExecutionContext(executionContext);
            self.addAllOnChange();
            self.FilterBTK(executionContext.getFormContext());
            self.enableDisableProposalWriterAssignedField(executionContext.getFormContext());
            self.enableDisableDateProposalCompletedfield(executionContext.getFormContext());
            self.showRiskExtension();
            self.showBusinessCre();
            self.setDefaultViewProposalWriter();
            self.checkIfmatterExist();
            //

            // self.leadGenandPursuitLeadOnChange();

            //  self.populateOppProduct(); 
            self.proposalTasks(); // stage on load is Identify
            //self.EngagementLetterAttachedOnChange(true);

            refreshRibbon(true);
        }
        catch (ex) {
            window.console.log("Error at CXP.Opportunity.OnLoad function: " + ex.message + "|" + "Stack: " + ex.stack);
            throw ex;
        }
    };
    //The onSave method
    self.onSave = function (executionContext) {
        setFormContextFromExecutionContext(executionContext);
        if (context.ui.getFormType() === 1) {
            self.createOpportunityName();


        }
        //window.setTimeout(function () {
        //    refreshRibbon(true);
        //}, 12);
        // self.setOppProductprice();

        self.proposalTasks();
        // self.setOppProductprice();
        //self.qualifyGetService();
        //self.identifyPhase();
        self.showRiskExtension();

        self.leadGenandPursuitLeadOnCreate();

    };



    //Functionality:
    //  If the opportunity isn't a parent then it will make sure the name is up to date. 
    //Trigger: OnSave
    /*	self.createOpportunityName = function ()
         {
               var currentName = getValue("name") || "";
               //var serviceRequestedText = getValue("cxp_servicerequested")[0].name;
               var accountText = getValue("parentaccountid")[0].name;
               var newName = "";
               if (currentName === "")
               {
                   newName = accountText + " - " + "Waiting for Service Line";
                   setValue("name", newName);
               }
         };
   
   */


    self.createOpportunityName = function () {
        debugger;

        var currentName = getValue("name") || "";
        //var serviceRequestedText = getValue("cxp_servicerequested")[0].name;
        var accountText = getValue("parentaccountid")[0].name;
        var givenName = getValue("cxp_opportunitygivenname");
        var newName;
        //if (currentName === "") {
        newName = accountText + " - " + givenName + " " + "Service";
        setValue("name", newName);
        // }
        //   else {
        //     setRequiredLevel("cxp_opportunitygivenname", "none");
        // }
    };


    //Functionality: Enable/Disable ProposalWriterAssigned based on the logged in userid profile
    //Trigger: onLoad of the page
    self.enableDisableProposalWriterAssignedField = function (formContext) {
        var context = Xrm.Utility.getGlobalContext();
        var loginUserID = context.userSettings.userId;
        //var ParentTab = getTab("Parent");
        // var ProposalRequestTab = getTab("tab_11");
        var result = JSON.parse(getSingle("systemusers", loginUserID, "cxp_receiveproposalneednotifications", false));
        if (result["cxp_receiveproposalneednotifications"]) {
            //setSectionVisibleFromTab(ProposalRequestTab, "Parent_section_8", true);
            setdisabled("cxp_defineproposalwriter", false);
        }
        else {
            //setSectionVisibleFromTab(ProposalRequestTab, "Parent_section_8", false);
            setdisabled("cxp_defineproposalwriter", true);
        }
    };
    //Functionality: Enable / Disable enableDisableDateProposalCompleted field based on the logged in userid profile
    //Trigger: onLoad of the page
    self.enableDisableDateProposalCompletedfield = function (formContext) {
        var context = Xrm.Utility.getGlobalContext();
        var loginUserID = context.userSettings.userId;
        var result = JSON.parse(getSingle("systemusers", loginUserID, "cxp_proposalwriter", false));
        if (result["cxp_proposalwriter"]) {
            //setSectionVisibleFromTab(ProposalRequestTab, "Parent_section_8", true);
            setdisabled("cxp_datesubmitted", false);
        }
        else {
            //setSectionVisibleFromTab(ProposalRequestTab, "Parent_section_8", false);
            setdisabled("cxp_datesubmitted", true);
        }
    };
    //Functionality: Shor Hide Risk Extension tab
    self.showRiskExtension = function () {

        var isLost = self.isLostOpp();
        var isFinalizePhase = self.InFinalizePhase();
        if (isFinalizePhase) {
            setTabVisible(getTab("Opportunity_Risk_Extension"), true); //shows Risk extension tab
            getTab("Opportunity_Risk_Extension").setFocus();
        }
        else if (isLost) { //If Opp is lost, keep Risk Ext Section Visible
            setTabVisible(getTab("Opportunity_Risk_Extension"), true); //shows Risk extension tab for closed as lost
        }
        else {
            setTabVisible(getTab("Opportunity_Risk_Extension"), false); // Risk extension tab
            setSectionVisibleFromTab(getTab("Opportunity_Risk_Extension"), "Matters", false)
        }
    };
    //Business Dev CRE section
    self.showBusinessCre = function () {
        var isCre = getValue("rg_gsmc");
        if (isCre) {
            setTabVisible(getTab("Business Development (CRE)"), true); //shows CRE tab
            setRequiredLevel("cpdc_assignedcre", "required");
            setRequiredLevel("cxp_validationsought", "required");
            setRequiredLevel("cxp_describehowthisopportunitywasidentified", "required");
            setRequiredLevel("cxp_describeyourrelationshiptotheclientthatha", "required");
            setRequiredLevel("cxp_whatworkdidyouprovidefromidentificationto", "required");
        }
        else {
            //hide CRE tab
            setRequiredLevel("cpdc_assignedcre", "none");
            setRequiredLevel("cxp_validationsought", "none");
            setRequiredLevel("cxp_describehowthisopportunitywasidentified", "none");
            setRequiredLevel("cxp_describeyourrelationshiptotheclientthatha", "none");
            setRequiredLevel("cxp_whatworkdidyouprovidefromidentificationto", "none");
            setTabVisible(getTab("Business Development (CRE)"), false);
        }
    };

    ////Functionality:In Qualify Phase/ ShoW Hide Service Areas Section / and decide if Prososal is required
    self.InGeneralRisk = function () {
        var inIdentify = self.inIdentifyPhase();
        var inQualify = self.inQualifyPhase();
        var inProposal1 = self.inProposalInProgressPhase();
        var inProposal2 = self.inProposalSubmitPhase();
        var inFinalize = self.InFinalizePhase();
        var InGeneralRisk = self.InGeneralRiskPhase();

        if (InGeneralRisk) {
            getTab("General and Required Opportunity Information").setFocus();

        }

    };

    ////Functionality:In Qualify Phase/ ShoW Hide Service Areas Section / and decide if Prososal is required
    self.qualifyGetService = function () {

        var inQualify = self.inQualifyPhase();
        var inProposal1 = self.inProposalInProgressPhase();
        var inProposal2 = self.inProposalSubmitPhase();
        var inFinalize = self.InFinalizePhase();
        var InGeneralRisk = self.InGeneralRiskPhase();
        if (inQualify) {
            //setVisible("cxp_proposalrequired", true);
            //setRequiredLevel("cxp_proposalrequired", "none");
            setRequiredLevel("pricelevelid", "required");
            getTab("tab_15").setFocus();


        }
        else if (inProposal1 || inProposal2 || inFinalize || InGeneralRisk) {
            setRequiredLevel("pricelevelid", "required");
            //setVisible("cxp_proposalrequired", true);
            //setRequiredLevel("cxp_proposalrequired", "required");
        }
        else {
            setRequiredLevel("pricelevelid", "none");

        }
    };

    ////Functionality: ShoW Hide Proposal Section


    self.proposalTasks = function () {
        debugger;

        var proposalreq = getValue("cxp_proposalrequired"); // Yes===280410000
        var inIdentify = self.inIdentifyPhase();
        var inQualify = self.inQualifyPhase();
        var inProposal1 = self.inProposalInProgressPhase();
        var inProposal2 = self.inProposalSubmitPhase();
        var inFinalize = self.InFinalizePhase();
        var InGeneralRiskPhase = self.InGeneralRiskPhase();


        var specialityServ = getValue("cxp_isthisacannabisrelatedopportunity");
        var independenceYN = getValue("cxp_care_afterconsiderationareweindependentof");
        var conflictsYN = getValue("cxp_carearethereanyconcernsaboutconflictofint");
        var distConflict = getValue("cxp_carewasafirmconflictofinterestdistributed");
        var assessConflict = getValue("evaluatefit");

        var isLost = self.isLostOpp();


        if (proposalreq === 280410001 || proposalreq === null || inIdentify || inQualify || InGeneralRiskPhase)  //proposal ==No
        {   //  No ===280410001
            setRequiredLevel("cxp_proposalrequired", "none");
            setRequiredLevel("cxp_dateproposalrequested", "none");
            setRequiredLevel("cxp_dateproposalrequested", "none");
            setRequiredLevel("cxp_dateproposaldue", "none");
            setRequiredLevel("rg_requesttype", "none");
            setRequiredLevel("quotecomments", "none");

            setTabVisible(getTab("tab_16"), true);

            setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_4", false);
            setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_3", false);
            setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_2", false)
            setTabVisible(getTab("tab_18"), false);
        }

        //  if(specialityServ != null && independenceYN != null && conflictsYN != null && distConflict != null && assessConflict != null) 
        //   {
        if (inProposal1) {
            getTab("tab_16").setFocus();
            setRequiredLevel("cxp_proposalrequired", "required");


            if (proposalreq === 280410000) {
                if (specialityServ != null && independenceYN != null && conflictsYN != null && distConflict != null && assessConflict != null) {

                    setRequiredLevel("cxp_dateproposalrequested", "required");
                    setRequiredLevel("cxp_dateproposaldue", "required");
                    setRequiredLevel("rg_requesttype", "required");
                    setRequiredLevel("quotecomments", "required");

                    setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_4", true);
                    setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_3", true);
                    setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_2", true)
                    setTabVisible(getTab("tab_18"), true);
                }
                else {

                    setRequiredLevel("cxp_dateproposalrequested", "none");
                    setRequiredLevel("cxp_dateproposaldue", "none");
                    setRequiredLevel("rg_requesttype", "none");
                    setRequiredLevel("quotecomments", "none");

                    setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_4", false);
                    setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_3", false);
                    setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_2", false)
                    setTabVisible(getTab("tab_18"), false);
                }

            }


        }
        if ((inFinalize || inProposal2) && proposalreq === 280410000)    // just show no focus
        {
            //setTabVisible(getTab("tab_16"), true);
            setRequiredLevel("cxp_proposalrequired", "required");
            setRequiredLevel("cxp_dateproposalrequested", "required");
            setRequiredLevel("cxp_dateproposaldue", "required");
            setRequiredLevel("rg_requesttype", "required");
            setRequiredLevel("quotecomments", "required");
            setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_4", true);
            setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_3", true);
            setSectionVisibleFromTab(getTab("tab_16"), "tab_16_section_2", true)
            setTabVisible(getTab("tab_18"), true);

        }




    };





    //Functionality: Tells if the BPF is in the finalize stage
    //Phase is Identify/Research

    //  Proposal Step 2  
    self.inIdentifyPhase = function () {
        var identifyStep = getValue("budgetstatus");
        if (identifyStep === 0) {
            return true;
        }
        else {
            return false;
        }
    };


    //Stage is Client Acceptance (Previously Finalize)
    self.InFinalizePhase = function () {
        var phase = getValue("budgetstatus");
        if (phase === 5) {
            return true;
        }
        else {
            return false;
        }
    };
    //Stage is qualified
    self.inQualifyPhase = function () {
        var qualifyStage = getValue("budgetstatus");
        if (qualifyStage === 1) {
            return true;
        }
        else {
            return false;
        }
    };
    //Stage is Proposal 1
    self.inProposalInProgressPhase = function () {
        var proposalStep1 = getValue("budgetstatus");
        if (proposalStep1 === 3) {
            return true;
        }
        else {
            return false;
        }
    };
    //  Proposal Step 2  
    self.inProposalSubmitPhase = function () {
        var proposalStep2 = getValue("budgetstatus");
        if (proposalStep2 === 4) {
            return true;
        }
        else {
            return false;
        }
    };
    //Opp Status Reason is Lost
    self.isLostOpp = function () {
        var statusLost = getValue("statecode");
        if (statusLost === 2) {
            return true;
        }
        else {
            return false;
        }
    };


    //Stage is General Risk
    self.InGeneralRiskPhase = function () {
        var riskPhase = getValue("budgetstatus");
        if (riskPhase === 2) {
            return true;
        }
        else {
            return false;
        }
    };
    //set estFees read only
    //self.estFeesReadOnly = function ()
    //{
    //var estFees = getValue("estimatedvalue");
    //estFees.setAttribute("readonly", true)
    //}



    //Functionality: Adds onChange callbacks to all of the fields that need it.
    //Trigger: addFieldInformation
    self.addAllOnChange = function () {
        //This on change event always needs to happen first
        addOnChange("cpdc_leadsourceid", self.ReferralContactOnChange);
        //    addOnChange("cxp_engagementletterattached", self.EngagementLetterAttachedOnChange);     
        addOnChange("cxp_legacysysstem", self.CannabisRelatedOpportunityOnChange);
        addOnChange("cxp_useaccountmailingaddress", self.PopulateAccountMailingAddress);
        addOnChange("parentaccountid", self.setDefaultValueofContact);
        //here//
        addOnChange("parentaccountid", self.setDefaultValueofBillingContact);
        addOnChange("budgetstatus", self.showRiskExtension);
        addOnChange("budgetstatus", self.InGeneralRisk);
        addOnChange("budgetstatus", self.qualifyGetService);
        addOnChange("budgetstatus", self.proposalTasks);
        addOnChange("cxp_proposalrequired", self.proposalTasks);
        addOnChange("cxp_proposalrequired", self.qualifyGetService);
        addOnChange("rg_gsmc", self.showBusinessCre);
        //	addOnChange("cxp_professionalpracticeleadersignature", self.saveForm);
        addOnChange("budgetstatus", self.saveForm);
        //   addOnChange("cxp_leadgenerator1",  self.leadGenandPursuitLeadOnChange);
        //   addOnChange("cxp_leadgenerator1",  self.saveForm);  
        addOnChange("createdon", self.saveForm);
        //addOnChange("cxp_opportunitygivenname", self.createOpportunityName);



    };


    //Functionality: force save
    self.saveForm = function () {
        context.data.entity.save();
    };
    /*
  //Functionality:
  //  Sets 'is this a cannabis related opportunity' to required and then
  //  sets the field to yes if it is and then shows an alert
  //Trigger: onChange of field 'cxp_isthisacannabisrelatedopportunity'
  self.CannabisRelatedOpportunityOnChange = function () {
      //var industry = getValue("cxp_industrypracticearea");
     
     // if (industry && industry[0].name.toLowerCase() === "cannabis") {
    
       //   setRequiredLevel("cxp_isthisacannabisrelatedopportunity", "required");
          //setValue("cxp_isthisacannabisrelatedopportunity", 280410001);
     // }
     // else if (industry && industry[0].name.toLowerCase() !== "cannabis") {
      //  setRequiredLevel("cxp_isthisacannabisrelatedopportunity", "not required");
     // }

      if (getValue("cxp_isthisacannabisrelatedopportunity") === 280410001) {

          var taxDirectorResult = JSON.parse(getMultiple("systemusers", "fullname", "cxp_cannabistaxpracticedirector eq true", false));

          var taxPartnerResult = JSON.parse(getMultiple("systemusers", "fullname", "cxp_cannabistaxpartner eq true", false));

          var alertStrings = {
              text:
                  "There is a special client Acceptance Process For Cannabis. A notification in this regard has been sent to " +
                  taxDirectorResult.value[0]["fullname"] +
                  " and " +
                  taxPartnerResult.value[0]["fullname"] +
                  ". Please discuss this opportunity with them prior to accepting this client."
          };
          var alertOptions = { height: 200, width: 450 };
          Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
      }
  };
*/
    self.CannabisRelatedOpportunityOnChange = function () {
        debugger;
        //280,410,003 Cannabis-Tax
        var alertOptions = {
            height: 200,
            width: 450
        };
        if (getValue("cxp_legacysysstem") === 280410003) {
            var taxDirectorResult = JSON.parse(getMultiple("systemusers", "fullname", "cxp_cannabistaxpracticedirector eq true", false));
            var taxPartnerResult = JSON.parse(getMultiple("systemusers", "fullname", "cxp_cannabistaxpartner eq true", false));
            var alertStrings = {
                text: "There is a special client Acceptance Process For Cannabis-Tax. A notification in this regard has been sent to " + taxDirectorResult.value[0]["fullname"] + " and " + taxPartnerResult.value[0]["fullname"] + ". Please discuss this opportunity with them prior to accepting this client."
            };
            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
        }
        // 280410000 Cannabis Assurance
        else if (getValue("cxp_legacysysstem") === 280410000) {
            var cannabisAssuranceResult = JSON.parse(getMultiple("systemusers", "fullname", "cxp_cannabisassuranceapprover eq true", false));
            var alertStrings2 = {
                text: "There is a special client Acceptance Process For Cannabis-Assurance. A notification in this regard has been sent to " + cannabisAssuranceResult.value[0]["fullname"] + ". Please discuss this opportunity with them prior to accepting this client."
            };
            //  var alertOptions = { height: 200, width: 450 };
            Xrm.Navigation.openAlertDialog(alertStrings2, alertOptions);
        }
    };
    //Functionality: Hides and shows fields based on engagement letter attached field
    //Trigger:onLoad and onChange of field 'cxp_engagementletterbypass'
    self.EngagementLetterAttachedOnChange = function (onLoad) {
        var no = 280410001;
        var alertStrings = {
            text: "Bypass request reason should explain why a signed Engagement Letter cannot be received prior to obtaining a matter number."
        };
        if (getValue("cxp_engagementletterattached") === no) {
            if (onLoad !== true) {
                Xrm.Navigation.openAlertDialog(alertStrings);
            }
            setVisible("cxp_elbypassreason", true);
            setValue("cxp_elbypassreason", null);
        }
        else {
            setVisible("cxp_elbypassreason", false);
            setValue("cxp_elbypassreason", null);
        }
    };
    //Functionality:  Sets the referral company when the referral company is set
    //Trigger:  onChange of field 'cpdc_leadsourceid'
    self.ReferralContactOnChange = function () {
        var referralContact = getValue("cpdc_leadsourceid");
        if (!referralContact) {
            return;
        }
        var successful = function (response) {
            var result = JSON.parse(response);
            if (!result["_parentcustomerid_value"]) {
                return;
            }
            var accountRef = new Array();
            accountRef[0] = new Object();
            accountRef[0].id = result["_parentcustomerid_value"];
            accountRef[0].entityType = result["_parentcustomerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
            accountRef[0].name = result["_parentcustomerid_value@OData.Community.Display.V1.FormattedValue"];
            setValue("cpdc_referralcompanyid", accountRef);
        };
        var error = function (error) {
            Xrm.Navigation.openAlertDialog(
                {
                    text: error
                });
        };
        getSingle("contacts", referralContact[0].id, "_parentcustomerid_value", true, successful, error);
    };
    // #region Filter BTK, APClerk, Contact
    //Functionality: It's functionality is to add presearch and remove presearch  for BTK.
    //Trigger: On every form load
    self.FilterBTK = function (formContext) {
        //context = executionContext.getFormContext();
        var control = formContext.getControl("cxp_billingtimekeeper");
        //Check if the control exist on the form
        if (control) {
            addPreSearch("cxp_billingtimekeeper", self.AddBTKFilter);
        }
        else {
            removePreSearch("cxp_billingtimekeeper", self.AddBTKFilter);
        }
    };
    //Functionality:  Filters the BillingTimekeeper before it is selected
    //Trigger:  PreSearch of field 'cxp_billingtimekeeper'
    self.AddBTKFilter = function () {
        var fetchQuery;
        //Build fetch
        fetchQuery = "<filter type='and'>" +
            "<condition attribute='cxp_billingtimekeeper' operator='eq' value='1' />" +
            "</filter>";
        //add custom filter
        addCustomFilter("cxp_billingtimekeeper", fetchQuery);
    };
    // #endregion
    // #region Set Default value of Contact
    //Functionality:Sets default value of the contact lookup when account field is updated
    //Triggers on change of account field.
    self.setDefaultValueofContact = function () {
        var account = getValue("parentaccountid");
        if (account !== null && account !== undefined && account.length > 0) {
            var accountID = account[0].id;
            var filter = "_parentcustomerid_value eq '" + accountID + "' and cxp_primarycontact eq true";
            var contactResult = JSON.parse(getMultiple("contacts", "fullname", filter, false));
            if (contactResult !== null && contactResult !== undefined && contactResult.value.length > 0) {
                var contact = contactResult.value[0];
                var contactName = contact.fullname;
                var contactID = contact.contactid;
                setLookupValue("parentcontactid", contactID, contactName, "contact", null);
            }
        }
    };
    //****/
    self.setDefaultValueofBillingContact = function () {
        var account = getValue("parentaccountid");
        if (account !== null && account !== undefined && account.length > 0) {
            var accountID = account[0].id;
            var filter = "_parentcustomerid_value eq '" + accountID + "' and cxp_primarycontact eq true";
            var contactResult = JSON.parse(getMultiple("contacts", "fullname", filter, false));
            if (contactResult !== null && contactResult !== undefined && contactResult.value.length > 0) {
                var contact = contactResult.value[0];
                var contactName = contact.fullname;
                var contactID = contact.contactid;
                setLookupValue("cxp_billingcontact", contactID, contactName, "contact", null);
            }
        }
    };
    // #endregion
    //#region address
    //Functionality: Gets the account address and populates on opportunity
    //Trigger: on update of the field "use account mailing address"
    self.PopulateAccountMailingAddress = function () {
        if (getValue("cxp_useaccountmailingaddress")) {
            var account = getValue("parentaccountid");
            var accountId = account ? account[0].id : "";
            var result = JSON.parse(getSingle("accounts", accountId, "address1_line1, address1_line2, address1_line3, address1_city, address1_postalcode, _cpdc_stateprovinceid_value, _cpdc_countryid_value", false));
            var address1_line1 = result["address1_line1"];
            var address1_line2 = result["address1_line2"];
            var address1_line3 = result["address1_line3"];
            var address1_city = result["address1_city"];
            var address1_postalcode = result["address1_postalcode"];
            if (result["_cpdc_countryid_value"] != null) {
                var countryRef = new Array();
                countryRef[0] = new Object();
                countryRef[0].id = result["_cpdc_countryid_value"];
                countryRef[0].entityType = result["_cpdc_countryid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                countryRef[0].name = result["_cpdc_countryid_value@OData.Community.Display.V1.FormattedValue"];
            }
            if (result["_cpdc_stateprovinceid_value"] != null) {
                var stateRef = new Array();
                stateRef[0] = new Object();
                stateRef[0].id = result["_cpdc_stateprovinceid_value"];
                stateRef[0].entityType = result["_cpdc_stateprovinceid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                stateRef[0].name = result["_cpdc_stateprovinceid_value@OData.Community.Display.V1.FormattedValue"];
            }
            setValue("cxp_billingstreet1", address1_line1);
            setValue("cxp_billingstreet2", address1_line2);
            setValue("cxp_billingstreet3", address1_line3);
            setValue("cxp_billingcity", address1_city);
            setValue("cxp_billingzippostalcode", address1_postalcode);
            setValue("cxp_billingcountryregion", countryRef);
            setValue("cxp_billingstateprovince", stateRef);
        }
        else {
            setValue("cxp_billingstreet1", null);
            setValue("cxp_billingstreet2", null);
            setValue("cxp_billingstreet3", null);
            setValue("cxp_billingcity", null);
            setValue("cxp_billingzippostalcode", null);
            setValue("cxp_billingcountryregion", null);
            setValue("cxp_billingstateprovince", null);
        }
    };
    // 

    self.setDefaultViewProposalWriter = function () {
        setLookupViewByName("cxp_defineproposalwriter", "Proposal Writers", true);
    }

    function setLookupViewByName(fieldName, viewName, asynchronous) {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/savedqueries?$select=savedqueryid&$filter=name eq '" + viewName + "'", asynchronous);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    if (results.value.length > 0) {
                        var savedqueryid = results.value[0]["savedqueryid"];
                        Xrm.Page.getControl(fieldName).setDefaultView(savedqueryid);
                    }
                    else {
                        Xrm.Utility.alertDialog(viewName + " view is not available.");
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    };


    self.EngagementLetterAttachedOnChange = function (onLoad) {
        var no = 280410001;
        var alertStrings = {
            text: "Bypass request reason should explain why a signed Engagement Letter cannot be received prior to obtaining a matter number."
        };
        if (getValue("cxp_engagementletterattached") === no) {
            if (onLoad !== true) {
                Xrm.Navigation.openAlertDialog(alertStrings);
            }
            setVisible("cxp_elbypassreason", true);
            setValue("cxp_elbypassreason", null);
        }
        else {
            setVisible("cxp_elbypassreason", false);
            setValue("cxp_elbypassreason", null);
        }
    };


    //set Opp to populate Opp product field if RE exist
    self.populateOppProduct = function (formContext) {

        var oppProd = getValue("cxp_opportunityproduct");
        var oppRE = getValue("cxp_riskextension")
        var oppREId = oppRE ? oppRE[0].id : "";
        oppREId = oppREId.replace("{", "").replace("}", "");
        var resultOppRE = JSON.parse(getSingle("cxp_opportunityriskextensions", oppREId, "_cxp_opportunityproduct_value ", false));
        var oppProductRE = resultOppRE["_cxp_opportunityproduct_value"];
        var _cxp_opportunityproduct_value_formatted = resultOppRE["_cxp_opportunityproduct_value@OData.Community.Display.V1.FormattedValue"];
        var _cxp_opportunityproduct_value_lookuplogicalname = resultOppRE["_cxp_opportunityproduct_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
        if (oppProd === null && oppRE !== null && oppRE !== "" && oppProductRE !== null && oppProductRE !== "") {
            setLookupValue("cxp_opportunityproduct", oppProductRE, _cxp_opportunityproduct_value_formatted, "opportunityproduct", null)
        }
    };
    //3940 if pursuit lead is diff than lead gen, create both connections wkf1 else create one connection with 1 role pursuit lead/lead gen
    //entityId is the id of the record for which the workflow will execute and workflowProcessId is the id of the workflow which we want to execute.

    /*    self.leadGenandPursuitLeadOnChange = function () { 
        
         var opportunityId = context.data.entity.getId();
          opportunityId = opportunityId.replace("{", "").replace("}", "");
   
    var entity = {
    "EntityId": opportunityId   }
    
     var WorkflowId0 = "e744c166-9d22-4ac7-8f08-b3f6a8128c1b";
     connectionWorkflow(WorkflowId0, entity);
     
      self.leadGenandPursuitLeadOnCreate();
   
   } */


    self.leadGenandPursuitLeadOnCreate = function () {

        debugger;

        // var WorkflowId0 = "e744c166-9d22-4ac7-8f08-b3f6a8128c1b";
        var WorkflowId1 = "55B88436-25C8-4667-9C0B-CE935F00370C";
        var WorkflowId2 = "b79f94bc-a1ae-4a35-b038-e184189f52bf";

        var pursuitLead = getValue("ownerid");
        var LeadGen = getValue("cxp_leadgenerator1");
        var opportunityId = context.data.entity.getId();
        opportunityId = opportunityId.replace("{", "").replace("}", "");

        var entity = {
            "EntityId": opportunityId
        }


        if (pursuitLead[0].name != null && LeadGen[0].name != null && pursuitLead[0].name != LeadGen[0].name && (getGridCount(getGrid("Pursuit_Team")) == 0)) {
            //connectionWorkflow(WorkflowId0, entity);
            connectionWorkflow(WorkflowId1, entity);
        }

        else if (pursuitLead[0].name != null && LeadGen[0].name != null && pursuitLead[0].name == LeadGen[0].name && (getGridCount(getGrid("Pursuit_Team")) == 0)) {
            //connectionWorkflow(WorkflowId0, entity);
            connectionWorkflow(WorkflowId2, entity);
        }

    };

    //
    function connectionWorkflow(WorkflowId, entity) {


        debugger;

        var req = new XMLHttpRequest();
        req.open("POST", parent.Xrm.Page.context.getClientUrl() + "/api/data/v9.0/workflows(" + WorkflowId + ")/Microsoft.Dynamics.CRM.ExecuteWorkflow", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;

                if (this.status === 200) {
                    Xrm.Utility.alertDialog("Success");
                } else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        //req.send(JSON.stringify(entity));
        req.send(JSON.stringify(entity));

    };


    //4335 If the client has a won opportunity then we can make it Existing. If not then it will be New ; cxp_neworexistingclient 

    //That field is a dropdown with only New or Existing,When the opportunity is opened then populate the New/Existing field with the above criteria

    self.checkIfmatterExist = function (formContext) {
        debugger;

        var client = getValue("parentaccountid")
        var clientId = client ? client[0].id : "";
        clientId = clientId.replace("{", "").replace("}", "");
        if (clientId != null && clientId != undefined) {
            //Xrm.WebApi.online.retrieveMultipleRecords("cpdc_matter", "?$filter=_cpdc_clientid_value eq").then(
            Xrm.WebApi.online.retrieveMultipleRecords("cpdc_matter", "?$filter= _cpdc_clientid_value eq " + clientId + "").then(
                function success(results) {
                    //for (var i = 0; i < results.entities.length; i++) {

                    //var cpdc_matterid = results.entities[i]["cpdc_matterid"];
                    if (results.entities.length > 0) {
                        setValue("cxp_neworexistingclient", 280410001, true) //existing  280,410,001
                    }
                    else if (results.entities.length == 0) {
                        setValue("cxp_neworexistingclient", 280410000, true) //new  280,410,000
                    }
                    //}
                },
                function (error) {
                    Xrm.Utility.alertDialog(error.message);
                }
            );
        }
    }

    //




    //#endregion
    return {
        onLoad: self.onLoad,
        onSave: self.onSave
    };
})();
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
