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
