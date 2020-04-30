if (typeof (CR) === "undefined") {
    CR = { _namespace: true };
}
if (typeof (CR.OpportunityRiskExtension) === "undefined") {
    CR.OpportunityRiskExtension = { _namespace: true };
}

CR.OpportunityRiskExtension.Form = (function () {
    "use strict";

    // #region Constants
    var notifications = {
        //closedOpps: "ClosedOpps",
		taxFields: "TaxFields",
		taxGFR: "TaxGFR",
        approvals: "Approvals",
        advisory: "Advisory",
		assurance: "Assurance",
		anyOther: "ExpectedMargin"

        // common: "Common",
        //accountSync: "accountSync",
        //billingAddress: "BillingAddress"
    };

    var approvalPairs = [
        ["cxp_engagementletterbypassapproval", "cxp_engagementletterbypasssignature"],
        ["cxp_partnerprincipalapprover", "cxp_partnerprincipalsignature"],
        ["cxp_ompapprover", "cxp_ompsignature"],
        ["cxp_riskreviewerapprover", "cxp_riskreviewersignature"],
        ["cxp_taxpracticedirectorapprover", "cxp_taxpracticedirectorsignature"],
        ["cxp_nationalleaderapprover", "cxp_nationalleadersignature"],
        ["cxp_cannabistaxdirectorapprover", "cxp_cannabistaxdirectorsignature"],
        ["cxp_professionalpracticeleaderapprover", "cxp_professionalpracticeleadersignature"],
        ["cxp_cannabistaxpartnerapprover", "cxp_cannabistaxpartnersignature"],
        ["cxp_engagementpartner", "cxp_engagementpartnersignature"]
    ];

    var self = {};

    var ServiceCategory = {
        Advisory: 280410000,
        Assurance: 280410001,
        Tax: 280410002,
        RCMS: 280410003
    };

    var ShortFormDesignator = {
        Tax: 280410000,
        Assurance: 280410001
    };

    var _serviceArea = {
        matterCode: "",
        serviceRequestedCategory: -1,
        isRestructuringAndLitigation: false
    };

    var _opp = {


    };



    //The onLoad method 
    self.onLoad = function (executionContext) {
        try {

            setFormContextFromExecutionContext(executionContext);
            self.GetServiceRequestedFields();
            //self.addFieldInformation();
            self.hideShowTabs();
            self.showNotifications();
            self.lockFieldsIfClosedAsWon();
            self.setDefaultValueofBillingContact();
            self.EngagementLetterAttachedOnChange(true);
            // self.showMatterDetailsGrid();
            self.setPhaseto6();
            self.setDefaultViewOTK();
            self.setDefaultViewSTK();
            self.setDefaultViewBTK();
			self.setDefaultViewMatterLocation();
			self.CheckApprovals();
			self.copyProductAndServicAreaOnLoad();			

        } catch (ex) {
            window.console.log("Error at CXP.OpportunityRiskExtension.OnLoad function: " + ex.message + "|" + "Stack: " + ex.stack);
            throw ex;
        }
    };

    //The onSave method
    self.onSave = function (executionContext) {
        self.hideShowTabs();
        self.showNotifications();
        self.lockFieldsIfClosedAsWon();
		self.matterYearOnChange();
		self.CheckApprovals();
    };


    //Functionality: Adds onChange callbacks to all of the fields that need it.
    self.addAllOnChange = function () {
        addOnChange("cxp_shortformreason", self.hideShowTabs);
        addOnChange("cxp_careassurancetierlevel", self.hideShowTabs);
        addOnChange("cxp_carehigherlevelofservicethanpreviouslyper", self.hideShowTabs);
        addOnChange("cxp_careindicatetheclientsrisklevel", self.hideShowTabs);
        addOnChange("cxp_shortformreason", self.showNotifications);
        addOnChange("cxp_careassurancetierlevel", self.showNotifications);
        addOnChange("cxp_carehigherlevelofservicethanpreviouslyper", self.showNotifications);
        addOnChange("cxp_careindicatetheclientsrisklevel", self.showNotifications);
        addOnChange("cxp_care_taxreturnyear", self.TaxReturnYearOnChange);
        addOnChange("cxp_fiscalyearbegin", self.FiscalYearBeginOnChange);
        addOnChange("new_care_fiscalyearend", self.FiscalYearEndOnChange);
        addOnChange("cxp_careadvisoryhasattestpartnerdeterminedind", self.AttestPartnerDeterminedIndependentOnChange);
        addOnChange("cxp_careadvisorynationalauditdeterminedindepe", self.AuditDeterminedIndependentOnChange);
        addOnChange("cxp_care_3sec", self.SecAdvisoryOnChange);
        addOnChange("cxp_care_3pcaob", self.PSCAOBAdvisoryOnChange);
        addOnChange("cxp_care_consideredtheeconomicfirmstandards", self.EconomonicFirmStandardsOnChange);
        addOnChange("cxp_matteryeartext", self.matterYearOnChange);
        addOnChange("cxp_engagementletterattached", self.EngagementLetterAttachedOnChange)
        //
        addOnChange("cxp_allowmovingtofinalizephase", self.requiredBeforeCloseAsWon);
        addOnChange("cxp_account", self.setDefaultValueofBillingContact);
        addOnChange("cxp_partnerprincipalsignature", self.saveForm);
        addOnChange("cxp_engagementpartnersignature", self.saveForm);
        addOnChange("cxp_cannabistaxdirectorsignature", self.saveForm);
        addOnChange("cxp_cannabistaxpartnersignature", self.saveForm);
        addOnChange("cxp_taxpracticedirectorsignature", self.saveForm);
        addOnChange("cxp_riskreviewersignature", self.saveForm);
        addOnChange("cxp_professionalpracticeleadersignature", self.saveForm);
        addOnChange("cxp_ompsignature", self.saveForm);
        addOnChange("cxp_engagementletterbypasssignature", self.saveForm);
        addOnChange("cxp_nationalleadersignature", self.saveForm);
        approvalPairs.forEach(function (pairs) {
            pairs.forEach(function (field) {
                addOnChange(field, self.CheckApprovals);
            });
        });

    };

    //Functionality: force save
    self.saveForm = function () {
        context.data.entity.save();
    };

    //Functionality: Adds PreSearchs and calls function to add onChanges. 
    //Trigger: onLoad
    /*	self.addFieldInformation = function () {
            addPreSearch("cxp_shortformreason", self.showServiceRequestedSelected);
    
            var modelmatterAtt = getAttribute("cxp_modelmatter");
            if (modelmatterAtt) {
                modelmatterAtt.controls.forEach(function (item) {
                    addPreSearch(item.getName(), self.filterModelMatter);
                });
            }
    
            var matterlocationAtt = getAttribute("cxp_matterlocation");
            if (matterlocationAtt) {
                matterlocationAtt.controls.forEach(function (item) {
                    addPreSearch(item.getName(), self.FilterMatterLocation);
                });
            }
            self.addAllOnChange();
        };
    
    */
    //Functionality: Gets servicearea Category and matterCode from the service area. 
    //Trigger: onLoad and onChange of field 'cxp_servicearea'
    self.GetServiceRequestedFields = function () {
        var id = getValue("cxp_servicearea");
        if (!id) {
            _serviceArea.matterCode = "";
            _serviceArea.serviceRequestedCategory = -1;
            return;
        }

        var result = JSON.parse(getSingle("products", id[0].id, "cxp_mattercode,cxp_servicecategory,cxp_restructuringlitigation,cxp_approvedforshortform,cxp_multipleservices", false));

        _serviceArea.matterCode = result["cxp_mattercode"];
        _serviceArea.serviceRequestedCategory = result["cxp_servicecategory"];
        _serviceArea.isRestructuringAndLitigation = result["cxp_restructuringlitigation"];
        _serviceArea.approvedForTaxShortForm = result["cxp_approvedforshortform"];
        _serviceArea.multipleServices = result["cxp_multipleservices"];
    };

    self.hideShowTabs = function () {

        //var General = getTab("General");

        try {
            if (_serviceArea.serviceRequestedCategory === ServiceCategory.Tax) {
                self.formLogicForTax();
            }
            if (_serviceArea.serviceRequestedCategory === ServiceCategory.Assurance) {
                self.formLogicForAssurance();
            }
            if (_serviceArea.serviceRequestedCategory === ServiceCategory.Advisory) {
                self.formLogicForAdvisory();
            }
            if (_serviceArea.serviceRequestedCategory === ServiceCategory.RCMS) {
                self.formLogicForRCMS();
            }
        } catch (ex) {
            window.console.log(ex);
        }

    };


    //***************** ADVISORY *************************/
    //Functionality: Hides/Shows sections and fields
    //Trigger: ShortFormLongFormLogic if service area is advisory
    self.formLogicForAdvisory = function () {
        self.hideAllTabs();
        setTabVisible(getTab("Risk Evaluation - Advisory"), true);   //shows advisory tab

        setValue("cxp_islongform", 280410000, true);   // sets - "Is long form" field to true
        setValue("cxp_formtype", 280410000, true);     // sets form type as Advisory

    };


    //******************** TAX *************************/
    //Functionality: Hides/Shows sections and fields
    //Trigger: ShortFormLongFormLogic if service area is tax
    self.formLogicForTax = function () {
        self.hideAllTabs(); 						 //hide all risk evaluation tabs
        setSectionVisibleFromTab(getTab("General"), "Tax_Short_Form_Qualifiers", true) // show short form qualifier section  for tax

        var shortFormPassed = false;
        var cannabisPassed = false;
        var account = getValue("cxp_account");
        var accountId = account ? account[0].id : "";
        var resultACC = JSON.parse(getSingle("accounts", accountId, "cxp_relatedclientnumber,accountnumber,_cxp_crindustrypracticearea1_value", false));
        var relatedClientNumber = resultACC["cxp_relatedclientnumber"];
        var accountNumber = resultACC["accountnumber"];
        var industryId = resultACC["_cxp_crindustrypracticearea1_value"];
        var shortFormReason = getValue("cxp_shortformreason");
        var clientRisk = getValue("cxp_careindicatetheclientsrisklevel");

        var shortFormDesignator = -1;
        if (shortFormReason) {
            var resultReasons = JSON.parse(getSingle("new_servicecodesfortaxshortreasonses", shortFormReason[0].id, "cxp_shortformdesignator", false));
            shortFormDesignator = resultReasons["cxp_shortformdesignator"];
        }

        var resultInd = JSON.parse(getSingle("cpdc_industrypracticeareas", industryId, "cpdc_name", false));
        var industryName = resultInd["cpdc_name"];

        var accountHasMattersWithTaxService = self.retrieveMattersForAccount(accountId, ServiceCategory.Tax);
        var parentAccount = JSON.parse(getSingle("accounts", accountId, "_parentaccountid_value", false));
        var parentAccountId = parentAccount["_parentaccountid_value"];
        var parentAccountHasMattersWithTaxService = self.retrieveMattersForAccount(parentAccountId, ServiceCategory.Tax);
        if (_serviceArea.approvedForTaxShortForm === false) {
            setValue("cxp_shortformreason", null);
        }

        var opp = getValue("cxp_opportunity");
        var oppId = opp ? opp[0].id : "";
        var resultOpp = JSON.parse(getSingle("opportunities", oppId, "cxp_isthisacannabisrelatedopportunity, estimatedvalue", false));
        var cannabisRelatedOpp = resultOpp["cxp_isthisacannabisrelatedopportunity"];
        var expectedFees = resultOpp["estimatedvalue"];

        // Rules below to decide if tax long form or tax short form tab should show up 
        // if it is a long form valudate restructuring check box as well to show/hide the restructuring section

        //Rule 1: If Cannabis || cannabis related opp || expected fee > 25000
        if ((industryId && industryName.toLowerCase() === "cannabis") || (cannabisRelatedOpp === 280410001) || (expectedFees > 25000)) {

            //Rule 1(i):	If account Industry = Cannabis  and service requested is a tax service 
            if (industryId && industryName.toLowerCase() === "cannabis") {
                setTabVisible(getTab("Risk Evaluation -Tax (Long Form)"), true);  // show tax long form tab
                if (_serviceArea.isRestructuringAndLitigation) {
                    setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", true);
                } else {
                    setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", false);
                }
                setValue("cxp_islongform", 280410000, true);  //sets - "Is long form" field to true
                setValue("cxp_formtype", 280410004, true);    // sets form type as Tax long form
                cannabisPassed = true;
            }
            //Rule 1(ii):If account Industry Not = Cannabis  and service requested is a tax service line, and the flag “Is this a cannabis related opportunity” = “yes” 

            if ((!industryId || industryName.toLowerCase() !== "cannabis") && cannabisRelatedOpp === 280410001) {
                setTabVisible(getTab("Risk Evaluation -Tax (Long Form)"), true);  // show tax long form tab
                if (_serviceArea.isRestructuringAndLitigation) {
                    setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", true);
                } else {
                    setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", false);
                }
                setValue("cxp_islongform", 280410000, true); //sets - "Is long form" field to true
                setValue("cxp_formtype", 280410004, true);   // sets form type as Tax long form
                cannabisPassed = true;
            }
        }
        else {
            //Rule 2 (i): If Account.AccountNumber is not null and the account already has a tax service line  (account has a matter already with tax service ) 
            //		  and  reason for short form contains data  and short form designator = tax and if service area approved for short form and Service Requested is a tax service line.
            if (accountNumber && accountHasMattersWithTaxService
                && shortFormReason
                && shortFormDesignator === ShortFormDesignator.Tax && _serviceArea.approvedForTaxShortForm === true) {
				setTabVisible(getTab("Tax_ShortForm"), true);   // show tax short form tab
                setValue("cxp_islongform", 280410001, true);  //sets - "Is long form" field to false
                setValue("cxp_formtype", 280410003, true);    // sets form type as Tax short
                shortFormPassed = true;
            }



            //Rule 2 (ii): If Account number is null and has a Parent Account with an account number and if the parent already has a tax service line(matter with tax service line) and reason for short form selected  
            //				and short form designator = tax and if service area approved for short form and Service Requested is a tax service line.

            if (!accountNumber && relatedClientNumber && parentAccountHasMattersWithTaxService
                && shortFormReason && shortFormDesignator === ShortFormDesignator.Tax && _serviceArea.approvedForTaxShortForm === true) {
                setTabVisible(getTab("Tax_ShortForm"), true)  // show tax short form tab
                setValue("cxp_islongform", 280410001, true); //sets - "Is long form" field to false
                setValue("cxp_formtype", 280410003, true);   // sets form type as Tax short
                shortFormPassed = true;
            }

            //Rule 2 (iii): If Short Reason Designator = Tax and “Indicate the Clients Risk Level” DOES = ‘C – Low OR NA’ and Expected fees < 20, 000 and account number not null
            //Changed to :If Short Reason Designator = Tax and Expected fees < 25, 000 and account number not null	

            if (shortFormDesignator === ShortFormDesignator.Tax &&
                includes(["280410003", "280410002"], clientRisk) &&
                ((expectedFees || 0) < 25000) &&
                accountNumber && _serviceArea.approvedForTaxShortForm === true) {
                setTabVisible(getTab("Tax_ShortForm"), true)  // show tax short form tab
                setValue("cxp_islongform", 280410001, true); //sets - "Is long form" field to false
                setValue("cxp_formtype", 280410003, true);   // sets form type as Tax short form
                shortFormPassed = true;
            }
        }

        //Rule 3: If Short Reason Designator is blank AND
        //		  If account Industry Not = Cannabis  and service requested is a tax service line, and the flag “Is this a cannabis related opportunity” = “no” or NULL
        if (!shortFormPassed && !cannabisPassed && (!industryId || industryName.toLowerCase() !== "cannabis") &&
            (!cannabisRelatedOpp || cannabisRelatedOpp === 280410000)) {
            setTabVisible(getTab("Risk Evaluation -Tax (Long Form)"), true);  // show tax long form tab
            if (_serviceArea.isRestructuringAndLitigation) {
                setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", true);
            } else {
                setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", false);
            }
            setValue("cxp_islongform", 280410000, true); //sets - "Is long form" field to true
            setValue("cxp_formtype", 280410004, true);   // sets form type as Tax long form
        }


    };

    //***************** ASSURANCE **********************/
    //Functionality: Hides/Shows sections and fields
    //Trigger: ShortFormLongFormLogic if service area is assurance
    self.formLogicForAssurance = function () {
        self.hideAllTabs();
        setSectionVisibleFromTab(getTab("General"), "Assurance_Short_Form_Qualifiers", true);

        var assuranceTierLevel = getValue("cxp_careassurancetierlevel");
        var higherLevelOfService = getValue("cxp_carehigherlevelofservicethanpreviouslyper");
        var shortFormPassed = false;
        var account = getValue("cxp_account");
        var accountId = account ? account[0].id : "";
        var resultACC = JSON.parse(getSingle("accounts", accountId, "cxp_relatedclientnumber,accountnumber", false));
        var relatedClientNumber = resultACC["cxp_relatedclientnumber"];
        var accountNumber = resultACC["accountnumber"];
        var shortFormReason = getValue("cxp_shortformreason");
        var shortFormDesignator = -1;
        if (shortFormReason) {
            var resultReasons = JSON.parse(getSingle("new_servicecodesfortaxshortreasonses", shortFormReason[0].id, "cxp_shortformdesignator", false));
            shortFormDesignator = resultReasons["cxp_shortformdesignator"];
        }
        var parentAccount = JSON.parse(getSingle("accounts", accountId, "_parentaccountid_value", false));
        var parentAccountId = parentAccount["_parentaccountid_value"];

        var accountHasMattersWithAssuranceService = self.retrieveMattersForAccount(accountId, ServiceCategory.Assurance);

        //Rule 1: if service area code is 2390 and the reason for short form is "lease accounting toolkit" then use short form else follow the rules 2-4 below
        if ((_serviceArea.matterCode === "2390") && (shortFormReason) && (shortFormReason[0].name.toLowerCase() === "lease accounting toolkit")) {
            setTabVisible(getTab("Assurance_ShortForm"), true);  // show assurance short form tab
            setValue("cxp_islongform", 280410001, true);        //sets - "Is long form" field to false    
            setValue("cxp_formtype", 280410006, true);          // sets form type as Assurance Short with Engagment Partner as Approver
            shortFormPassed = true;
        }
        else {

            //Rule 2:If account.accountnumber is null Account.cxp_RelatedClientNumber is NOT null
            //		 And Opportunity.cxp_careassurancetierlevel = ‘III’ and Service Requested is an Assurance Service line
            if (accountNumber && accountHasMattersWithAssuranceService && (assuranceTierLevel === 280410002) && (higherLevelOfService === 280410001)) {
                setTabVisible(getTab("Assurance_ShortForm"), true);  // show assurance short form tab
                setValue("cxp_islongform", 280410001, true);        //sets - "Is long form" field to false    
                setValue("cxp_formtype", 280410001, true);          // sets form type as Assurance Short
                shortFormPassed = true;
            }

            //Rule 3: If Account.AccountNumber is not null and the account already has an assurance service line  (account has a matter with assurance service) 
            // and (Tier Rating = 3 OR Higher level of Service = No) and Service Requested is an Assurance Service line
            //------ Rule updated recently to look at parent account/ other child accounts as well   (VSTS 3643)
            var parentAccountHasMattersWithAssurance = self.retrieveMattersForAccount(parentAccountId, ServiceCategory.Tax);
            var accountwithSameParentHaveMatterWithAssurance = self.retrieveMattersForAccountsWithSameParent(ServiceCategory.Assurance);

            if (accountNumber && (accountHasMattersWithAssuranceService || parentAccountHasMattersWithAssurance || accountwithSameParentHaveMatterWithAssurance) && ((assuranceTierLevel === 280410002) && (higherLevelOfService === 280410001))) {
                setTabVisible(getTab("Assurance_ShortForm"), true);  // show assurance short form tab
                setValue("cxp_islongform", 280410001, true);  //sets - "Is long form" field to false    
                setValue("cxp_formtype", 280410001, true);    // sets form type as Assurance short
                shortFormPassed = true;
            }
            //Rule 4: If the above rules fail and the service requested is assurance Service line( if the short form qualifying questions fail )
            if (!shortFormPassed) {
                setTabVisible(getTab("Risk Evaluation - Assurance (Long Form)"), true);
                if (_serviceArea.isRestructuringAndLitigation) {
                    setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", true);
                } else {
                    setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", false);
                }
                setValue("cxp_islongform", 280410000, true); //sets - "Is long form" field to true    
                setValue("cxp_formtype", 280410002, true);   // sets form type as Assurance long form
            }
        }

    };


    //******************* RCMS ***********************/
    //Functionality: Hides/Shows sections and fiels
    //Trigger:ShortFormLongFormLogic if service area is RCMS
    self.formLogicForRCMS = function () {
        self.hideAllTabs();
        setTabVisible(getTab("RCMS"), true)
        setValue("cxp_islongform", 280410000, true); //sets - "Is long form" field to true    
        setValue("cxp_formtype", 280410005, true);   // sets form type as RCMS

    };


    // Hide all Risk evaluation tabs
    self.hideAllTabs = function () {
        setTabVisible(getTab("Risk Evaluation - Assurance (Long Form)"), false);
        setTabVisible(getTab("Risk Evaluation - Advisory"), false);
        setTabVisible(getTab("Risk Evaluation -Tax (Long Form)"), false);
        setTabVisible(getTab("Assurance_ShortForm"), false);
        setTabVisible(getTab("Tax_ShortForm"), false);
        setTabVisible(getTab("Risk Evaluation - Advisory"), false);
        setTabVisible(getTab("RCMS"), false);

        setSectionVisibleFromTab(getTab("General"), "Assurance_Short_Form_Qualifiers", false);
        setSectionVisibleFromTab(getTab("General"), "Tax_Short_Form_Qualifiers", false);
        setSectionVisibleFromTab(getTab("General"), "Restruction_and_Litigation", false);
    };


    // Retrieve matters for the account/parent account
    self.retrieveMattersForAccount = function (accountId, serviceCategory) {
        if (accountId !== null) {
            var fetchXML = "<fetch>" +
                "<entity name='cpdc_matter' >" +
                "<attribute name='cpdc_name' />" +
                "<attribute name='statuscode' />" +
                "<attribute name='cxp_opportunityproduct' />" +
                "<attribute name='cpdc_matternumber' />" +
                "<attribute name='cpdc_clientid' />" +
                "<attribute name='cpdc_matterid' />" +
                "<filter type='and'>" +
                "<condition attribute='statecode' operator='eq' value='0' />" +
                "<condition attribute='cpdc_clientid' operator='eq' value='" + removeBraces(accountId) + "' />" +
                "<condition attribute='cpdc_matternumber' operator='not-null' />" +
                "</filter>" +
                "<link-entity name='product' alias='ad' link-type='inner' to='cxp_product' from='productid'>" +
                "<filter type='and'>" +
                "condition attribute='cxp_servicecategory' operator='eq' value='" + serviceCategory + "' />" +
                "</filter>" +
                "</link-entity>" +

                "</entity>" +
                "</fetch>";
            var response = getFromFetch("cpdc_matters", fetchXML, false);

            var result = JSON.parse(response);

            //Returning greater than 0 because atleast one matter
            return result.value.length > 0;
        }
        else {
            return false;
        }
    };


    //Functionality:  Checks to see if there are any other Accounts  with the same "Parent account"
    //  that have a have service area and service area category is the type passed, in status won
    //Trigger: Can be called from any method
    self.retrieveMattersForAccountsWithSameParent = function (serviceCategory) {
        var account = getValue("parentaccountid");
        var accountId = account ? account[0].id : "";
        var parentAccount = JSON.parse(getSingle("accounts", accountId, "_parentaccountid_value", false));

        var parentAccountId = parentAccount["_parentaccountid_value"];

        var relatedClientNumber = getValue("cxp_relatedclientnumber");

        if (relatedClientNumber !== null) {
            var fetchXML = "<fetch>" +
                "<entity name='cpdc_matter' >" +
                "<attribute name='cpdc_name' />" +
                "<attribute name='statuscode' />" +
                "<attribute name='cxp_servicerequested' />" +
                "<attribute name='cpdc_matternumber' />" +
                "<attribute name='cpdc_clientid' />" +
                "<attribute name='cpdc_matterid' />" +
                "<filter type='and'>" +
                "<condition attribute='statecode' operator='eq' value='0' />" +
                "<condition attribute='cpdc_matternumber' operator='not-null' />" +
                "<condition attribute='cxp_relatedclientnumber' operator='eq' value='" + relatedClientNumber + "' />" +
                "</filter>" +
                "<link-entity name='rg_servicearea' from='rg_serviceareaid' to='cxp_servicerequested' link-type='inner' alias='ServiceArea' >" +
                "<filter type='and'>" +
                "<condition attribute='cxp_servicecategory' operator='eq' value='" + serviceCategory + "' />" +
                "</filter>" +
                "</link-entity>" +
                "</entity>" +
                "</fetch >";
            var response = getFromFetch("cpdc_matters", fetchXML, false);

            var result = JSON.parse(response);

            //Returning greater than 0 because atleast one matter
            return result.value.length > 0;
        }
        else {
            return false;
        }

    };

    // Functionality: Shows notifications of the fields that needs to be updated 
    // Trigger: on load,  on Save
    self.showNotifications = function () {
        // Code to validate form type and then call the functions to validate if required fields populated accordingly; and then set the field "Allow moving to finalize phase" 
        // If short form or RCMS form, set this field "Allow moving to finalize phase" to "Yes"
        // this new field "Allow moving to finalize phase" is added to BPF and hidden section and this field is a 2 option set and marked mandatory
        // So unless this field is checked as "Yes", we cannot proceed to Finalize phase, and this is a ready only field , to not allow users to edit it

		if ((getValue("cxp_islongform") === 280410001) || (getValue("cxp_formtype") === 280410005)) // if short form or RCMS long form
		{
			setValue("cxp_allowmovingtofinalizephase", 1, true); // sets the field "Allow moving to finalize phase" as Yes

			context.ui.clearFormNotification(notifications.assurance);
			context.ui.clearFormNotification(notifications.taxFields);
			context.ui.clearFormNotification(notifications.advisory);
			context.ui.clearFormNotification(notifications.anyOther);
			context.ui.clearFormNotification(notifications.taxGFR);

			if (getValue("cxp_islongform") === 280410001) { //short form
				var closedAsWon = getValue("cxp_closedaswon");
			    var closedAsLost = getValue("cxp_closedaslost");
			    var approvalsRequested = getValue("cxp_approvalsrequested");
				if ((closedAsWon) || (closedAsLost) || (approvalsRequested)) {
					context.ui.clearFormNotification(notifications.anyOther);
					context.ui.clearFormNotification(notifications.taxGFR);
				}
				else {
					if (_serviceArea.matterCode === "1040") {
						var account = getValue("cxp_account");
						var accountId = account ? account[0].id : "";
						var resultACC = JSON.parse(getSingle("accounts", accountId, "cxp_gfr_client_name", false));
						var GFR = resultACC["cxp_gfr_client_name"];
						if ((GFR === "") || (GFR === null)) {
							setValue("cxp_allowmovingtofinalizephase", 0, true); // sets the field "Allow moving to finalize phase" as No
							context.ui.setFormNotification(
								"You must have a GFR Name on the account to do a 1040 for this client",
								"WARNING", notifications.taxGFR);
						}
						else {
							context.ui.clearFormNotification(notifications.taxGFR);
						}
					}
					if (getValue("cxp_expectedprofitability") === null) {         // 
						context.ui.setFormNotification(
							"Under Fees section, Expected Margin % field should be updated to proceed to request approvals",
							"WARNING", notifications.anyOther);
					}
					else {
						context.ui.clearFormNotification(notifications.anyOther);
					}
				}
			}
		

        }
        else {
            if (getValue("cxp_formtype") === 280410000) { // if Advisory Long form

                var advisoryFieldsPopulated = self.AreAdvisoryNOIFieldsPopulated();
                if (advisoryFieldsPopulated) {
					setValue("cxp_allowmovingtofinalizephase", 1, true); // sets the field "Allow moving to finalize phase" as Yes
					self.areOtherFieldsPopulated();				

                }
                else {
					setValue("cxp_allowmovingtofinalizephase", 0, true); // sets the field "Allow moving to finalize phase" as No
                }
            }
            if (getValue("cxp_formtype") === 280410004) { // if Tax Long form
                var taxFieldsPopulated = self.AreTaxFieldsPopulated();
                if (taxFieldsPopulated) {
					setValue("cxp_allowmovingtofinalizephase", 1, true); // sets the field "Allow moving to finalize phase" as Yes
					self.areOtherFieldsPopulated();
                }
                else {
					setValue("cxp_allowmovingtofinalizephase", 0, true); // sets the field "Allow moving to finalize phase" as No
                }
            }
            if (getValue("cxp_formtype") === 280410002) { // if Assurance Long form
                var assuranceFieldsPopulated = self.AreAssuranceFieldsPopulated();
                if (assuranceFieldsPopulated) {
					setValue("cxp_allowmovingtofinalizephase", 1, true);  // sets the field "Allow moving to finalize phase" as Yes
					self.areOtherFieldsPopulated();
                }
                else {
					setValue("cxp_allowmovingtofinalizephase", 0, true);  // sets the field "Allow moving to finalize phase" as No
                }
            }

        }

    };

    //Functionality: If assurance long form make sure fields are filled out propertly
    //Trigger: Called from the function "showNotifications"
    self.AreAssuranceFieldsPopulated = function () {
        var assuranceLongFormTab = getTab("Risk Evaluation - Assurance (Long Form)");
        var returnValue = true;
        var listAssuranceFields = "";
        var semicolon = "; ";

        if (!getTabVisible(assuranceLongFormTab)) {
            context.ui.clearFormNotification(notifications.assurance);
            return true;
        }

        if (!getValue("cxp_careaudityear")) {
            returnValue = false;
            listAssuranceFields = listAssuranceFields + getLabel("cxp_careaudityear") + semicolon;
        }
        if (!getValue("cxp_careassurancetierlevel")) {
            returnValue = false;
            listAssuranceFields = listAssuranceFields + getLabel("cxp_careassurancetierlevel") + semicolon;
        }
        if (getValue("cxp_careassurancetierlevel") === 280410000) {  // if assurance tier level is Tier 1
            if (!getValue("cxp_designatedreviewer")) {
                returnValue = false;
                listAssuranceFields = listAssuranceFields + getLabel("cxp_designatedreviewer") + semicolon;
            }
            if (!getValue("cxp_caredidyouinformtherevieweroftheselection")) {
                returnValue = false;
                listAssuranceFields = listAssuranceFields + getLabel("cxp_caredidyouinformtherevieweroftheselection") + semicolon;
            }
        }

        if (returnValue) {
            context.ui.clearFormNotification(notifications.assurance);
        } else {
            context.ui.setFormNotification(
                "Under Risk Evaluation-Assurance section, following fields should be updated, to proceed: " + listAssuranceFields,
                "WARNING", notifications.assurance);
        }

        return returnValue;
    };

    //Functionality: If tax long form make sure fields are filled out propertly
    //Trigger: Called from the function "showNotifications"
	self.AreTaxFieldsPopulated = function () {
        var yes = 280410000;
        var no = 280410001;
		var taxLongFormTab = getTab("Risk Evaluation -Tax (Long Form)");
        var returnValue = true;
        var listTaxFields = "";
        var semicolon = "; ";
		var describe = "-Describe";

        if (!getTabVisible(taxLongFormTab)) {
            context.ui.clearFormNotification(notifications.taxFields);
            return true;
        }
        if (!getValue("cxp_care_unusualriskscausenotacceptclient")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_care_unusualriskscausenotacceptclient") + semicolon;
        }
        else if (getValue("cxp_care_unusualriskscausenotacceptclient") === yes && !getValue("cxp_care_unusualriskscausenotacceptclienttext")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_care_unusualriskscausenotacceptclient") + describe + semicolon;
        }
        if (!getValue("cxp_careistheclientundergoingataxexamination")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_careistheclientundergoingataxexamination") + semicolon;
        }
        else if (getValue("cxp_careistheclientundergoingataxexamination") === yes && !getValue("cxp_caretaxexaminationspecifics")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_caretaxexaminationspecifics") + semicolon;
        }
        if (!getValue("cxp_care_expertisetoperformtheengagement")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_care_expertisetoperformtheengagement") + semicolon;
        }
        else if (getValue("cxp_care_expertisetoperformtheengagement") === no && !getValue("cxp_care_expertisetoperformtheengagementtext")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_care_expertisetoperformtheengagement") + describe + semicolon;
        }
        if (!getValue("cxp_care_staffingcommitmentwithinourcapacity")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_care_staffingcommitmentwithinourcapacity") + semicolon;
        }
        else if (getValue("cxp_care_staffingcommitmentwithinourcapacity") === yes && !getValue("cxp_care_staffingcommitmentwithincapacitytext")) {
            returnValue = false;
            listTaxFields = listTaxFields + getLabel("cxp_care_staffingcommitmentwithinourcapacity") + describe + semicolon;
        }
        if (returnValue) {
            context.ui.clearFormNotification(notifications.taxFields);
        } else {
            context.ui.setFormNotification(
                "Under Risk Evaluation-Tax section, following fields should be updated, to proceed: " + listTaxFields,
                "WARNING", notifications.taxFields);
		}

		if (_serviceArea.matterCode === "1040") {
			var account = getValue("cxp_account");
			var accountId = account ? account[0].id : "";
			var resultACC = JSON.parse(getSingle("accounts", accountId, "cxp_gfr_client_name", false));
			var GFR = resultACC["cxp_gfr_client_name"];

			if ((GFR === "") || (GFR === null)) {
				returnValue = false;
				context.ui.setFormNotification(
					"You must have a GFR Name on the account to do a 1040 for this client.",
					"WARNING", notifications.taxGFR);
			}
			else {
				context.ui.clearFormNotification(notifications.taxGFR);
			}
		}
			
        return returnValue;
    };

    //Functionality:  If Advisory long form make sure fields are filled out propertly
    //Trigger: Called from the function "showNotifications"
    self.AreAdvisoryNOIFieldsPopulated = function () {
        var advisoryTab = getTab("Risk Evaluation - Advisory");
        var attestAreaGrid = getGrid("AdvisoryAttestServicesWePerform");
        var returnValue = true;
        var listAdvisoryFields = "";
        var semicolon = "; ";
        var gridLabel = "Attest Service Areas"

        if (!getTabVisible(advisoryTab)) {
            context.ui.clearFormNotification(notifications.advisory);
            return true;
        }
        var gridCount = getGridCount(attestAreaGrid);
        if (gridCount === -1) {
            window.setTimeout(function () {
                context.ui.refreshRibbon();
            },
                3000);
            return false;
        }
        if (!getValue("cxp_careisthisentityanattestclientofthefirm")) {
            returnValue = false;
            listAdvisoryFields = listAdvisoryFields + getLabel("cxp_careisthisentityanattestclientofthefirm") + semicolon;
        }
        else if (getValue("cxp_careisthisentityanattestclientofthefirm") === 280410000) {
            if (gridCount === 0) {
                returnValue = false;
                listAdvisoryFields = listAdvisoryFields + gridLabel + semicolon;
            }
            if (!(getValue("cxp_care_3dol") || getValue("cxp_care_3gaas") || getValue("cxp_care_3gagas") || getValue("cxp_care_3pcaob") || getValue("cxp_care_3sec"))) {
                returnValue = false;
                listAdvisoryFields = listAdvisoryFields + getLabel("cxp_care_3dol") + "/" + getLabel("cxp_care_3gaas") + "/" + getLabel("cxp_care_3gagas") + "/" + getLabel("cxp_care_3pcaob") + "/" + getLabel("cxp_care_3sec") + "(one of these has to be Yes)" + semicolon;
            }
        }
        if (returnValue) {
            context.ui.clearFormNotification(notifications.advisory);
        } else {
            context.ui.setFormNotification(
                "Under Risk Evaluation-Advisory section, following fields should be updated, to proceed: " + listAdvisoryFields,
                "WARNING",
                notifications.advisory);
        }
        return returnValue;

    };

	self.areOtherFieldsPopulated = function () {
		var feesTab = getTab("FinalizeTab");
		//var oppApproved = getValue("cxp_opportunityapproved");
		var closedAsWon = getValue("cxp_closedaswon");
		var closedAsLost = getValue("cxp_closedaslost");
		var approvalsRequested = getValue("cxp_approvalsrequested");
		//var oppApproved = getValue("cxp_opportunityapproved");
		var oppApproved = self.isOpportunityApproved();
		//Xrm.Navigation.openAlertDialog (approvalsRequested);
		if ((!getTabVisible(feesTab)) || (closedAsWon) || (closedAsLost) || (approvalsRequested)) {
			context.ui.clearFormNotification(notifications.anyOther);
		}
		else {
			if (getValue("cxp_expectedprofitability") === null) {
				context.ui.setFormNotification(
					"Under Fees section, Expected Margin % field should be updated to proceed to request approvals",
					"WARNING", notifications.anyOther);
			}
			else {
				context.ui.clearFormNotification(notifications.anyOther);
			}
		}
		
	};

    ////Functionality:Sets a filter on field 'cxp_shortformdesignator' based on the selected Service Requested
    ////Trigger: Called from serviceRequestedOnChange function
    self.showServiceRequestedSelected = function () {
        var filter = "";
        var attribute = getValue("cxp_opportunityproduct");
        if (attribute) {
            var serviceCategory = _serviceArea.serviceRequestedCategory;
            var approvedForTaxShort = _serviceArea.approvedForTaxShortForm;
            if (serviceCategory === 280410001) {
                filter = "<filter type='and'><condition attribute='cxp_shortformdesignator' operator='eq' value='280410001' /></filter>";
            }
            else if (serviceCategory === 280410002 && approvedForTaxShort === true) {
                filter = "<filter type='and'><condition attribute='cxp_shortformdesignator' operator='eq' value='280410000' /></filter>";
            }
            else {
                filter = "<filter type='and'><condition attribute='createdon' operator='null' /></filter >";
            }
            addCustomFilter("cxp_shortformreason", filter);
        } else {
            addCustomFilter("cxp_shortformreason", "");
        }
    };

    //Functionality: Filters the locations before it is selected
    //Trigger:  PreSearch of field 'cxp_matterlocation'
    self.FilterMatterLocation = function () {
        var filter = "<filter type='and'>" +
            "<condition attribute='statecode' operator='eq' value='0' />" +
            "<filter type='or'>" +
            "<condition attribute='cxp_physicaloffice' operator='eq' value='1' />" +
            "<condition attribute='rg_officeid' operator='eq' uiname='Non Folder Set Up Required' uitype='rg_office' value='{BDF3C417-619D-E711-8122-E0071B669EB1}' />" +
            "</filter>" +
            "<condition attribute='rg_officeid' operator='ne' uiname='Cayman Islands' uitype='rg_office' value='{92D5004A-BE6C-E711-8123-E0071B6A71A1}' />" +
            "<condition attribute='rg_officeid' operator='ne' uiname='Woodland Hills' uitype='rg_office' value='{EC14D07E-648E-E711-8123-E0071B6A50A1}' />" +
            "<condition attribute='rg_officeid' operator='ne' uiname='Annapolis' uitype='rg_office' value='{605CE905-BE69-E411-993F-005056BE3808}' />" +
            "</filter>";

        getAttribute("cxp_matterlocation").controls.forEach(function (item) {
            addCustomFilter(item.getName(), filter);
        });
    };

    //Functionality: Filters the modelmatter before it is selected
    //Trigger: PreSearch of field 'cxp_modelmatter'
    self.filterModelMatter = function () {
        var departmentcode = getValue("cxp_departmentcode");
        var fetchQuery;
        var account = getValue("cxp_account");
        var result = JSON.parse(getSingle("accounts", account[0].id, "accountnumber,cxp_relatedclientnumber", false));
        var accountnumber = result["accountnumber"];
        var relatedclientnumber = result["cxp_relatedclientnumber"];
        if (departmentcode !== null && departmentcode !== undefined) {
            if (relatedclientnumber !== null && relatedclientnumber !== undefined) {
                fetchQuery = "<filter type='and'>" +
                    "<condition attribute='cxp_relatedclientnumber' operator='eq' value='" + relatedclientnumber + "' />" +
                    "<condition attribute='cxp_departmentcode' operator='eq' value='" + departmentcode + "' />" +
                    "</filter>";
            }
            else if (accountnumber !== null && accountnumber !== undefined) {
                fetchQuery = "<filter type='and'>" +
                    "<condition attribute='cxp_clientnumber' operator='eq' value='" + accountnumber + "' />" +
                    "<condition attribute='cxp_departmentcode' operator='eq' value='" + departmentcode + "' />" +
                    "</filter>";
            }
        }
        else {
            if (relatedclientnumber !== null && relatedclientnumber !== undefined) {
                fetchQuery = "<filter type='and'>" +
                    "<condition attribute='cxp_relatedclientnumber' operator='eq' value='" + relatedclientnumber + "' />" +
                    "</filter>";
            }
            else if (accountnumber !== null && accountnumber !== undefined) {
                fetchQuery = "<filter type='and'>" +
                    "<condition attribute='cxp_clientnumber' operator='eq' value='" + accountnumber + "' />" +
                    "</filter>";
            }
        }

        getAttribute("cxp_modelmatter").controls.forEach(function (item) {
            addCustomFilter(item.getName(), fetchQuery);
        });
    };

    //Functionality:  If the fiscal year end has a value than call setMatterYear
    //Trigger:  onChange of field 'cxp_fiscalyearend'
    self.FiscalYearEndOnChange = function () {
        var tax = 280410001;
        var fiscalEnd = getValue("cxp_fiscalyearend");

        if (fiscalEnd && getValue("cxp_departmentcode") === tax) {
            self.SetMatterYear();
        }
    };

    //Functionality: If the fiscal year begin has a value than call setMatterYear
    //Trigger:onChange of field 'cxp_fiscalyearbegin'
    self.FiscalYearBeginOnChange = function () {
        var fiscalBegin = getValue("cxp_fiscalyearbegin");
        if (fiscalBegin) {
            self.SetMatterYear();
        }
    };
    //Functionality: If the tax return year has a value than call setMatterYear
    //Trigger: onChange of field 'cxp_care_taxreturnyear'
    self.TaxReturnYearOnChange = function () {
        var returnYear = getValue("cxp_care_taxreturnyear");

        if (returnYear) {
            self.SetMatterYear();
        }
    };


    //Functionality:Sets the matter year according to logic including fiscal year end,fiscal year begin, and matter codes
    //Trigger: Can be called from any method
    self.SetMatterYear = function () {
        var tax = 280410001;
        var dec31St = 279640011;
        var isComplianceMatter = self.IsComplianceMatter();
        var fiscalYearBegin = getValue("cxp_fiscalyearbegin");
        var fiscalYearEnd = getValue("cxp_fiscalyearend");

        if (getValue("cxp_departmentcode") === tax && fiscalYearEnd !== dec31St) {
            setVisible("cxp_fiscalyearbegin");
        } else {
            setValue("cxp_fiscalyearbegin", null);
            setVisible("cxp_fiscalyearbegin", false);
        }

        if (_serviceArea.matterCode === "0709") {
            self.SetMatterYear00();
        } else if (isComplianceMatter && fiscalYearEnd === dec31St) {
            self.SetMatterYearToTaxReturnYear();
        } else if (isComplianceMatter && fiscalYearEnd !== dec31St &&
            (fiscalYearBegin === null || fiscalYearBegin === "")) {
            var alertStrings = {
                text: "Please input the year in which the fiscal year starts - YYYY."
            };
            Xrm.Navigation.openAlertDialog(alertStrings).then(function () {
                setFocus("cxp_fiscalyearbegin");
            },
                function () {
                    setFocus("cxp_fiscalyearbegin");
                });
        } else if (isComplianceMatter && fiscalYearEnd !== dec31St && fiscalYearBegin) {
            if (fiscalYearBegin.length === 2) {
                setValue("cxp_matteryeartext", fiscalYearBegin);
            } else if (fiscalYearBegin.length > 2) {
                setValue("cxp_matteryeartext", fiscalYearBegin.slice(-2));
            }
        } else if (includes(["1040", "1041", "0706"], _serviceArea.matterCode)) {
            if (getValue("cxp_matteryeartext") !== "00") {
                var confirmStrings = { text: "Do you want to set up a '00' matter?", title: "Set Matter Year", confirmButtonLabel: "Yes", cancelButtonLabel: "No" };
                var confirmOptions = { height: 200, width: 450 };
                Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                    function (success) {
                        if (success.confirmed) {
                            setRequiredLevel("cxp_care_taxreturnyear", "none");
                            self.SetMatterYear00();
                        } else {
                            setRequiredLevel("cxp_care_taxreturnyear", "required");
                            self.SetMatterYearToTaxReturnYear();
                        }
                    });
            }
        } else if (self.IsConsultingMatter()) {
            self.SetMatterYearToTaxReturnYear();
        }
    };


    //Functionality:Pass a matter code to the method and it will tell you if the matter is a compliance matter
    //Trigger: Called from another method.
    self.IsComplianceMatter = function () {
        var complianceMatters = [
            "1099", "PAYR", "PROP", "SALE", "1395", "0990",
            "1065", "10NR", "5500", "55EZ", "99PF", "CCON",
            "CORP", "CPAR", "CSUB", "INTL", "QSSS", "SCOR",
            "SPAR"
        ];
        return includes(complianceMatters, _serviceArea.matterCode);
    };

    //Functionality: Pass a matter code to the method and it will tell you if the matter is a consulting matter
    //Trigger: Called from another method.
    self.IsConsultingMatter = function () {
        var consultingMatters = [
            "FIDU", "1000", "1001", "1002", "1003", "1004", "1005", "1007",
            "1008", "1009", "1010", "1011", "1026", "1046", "1047", "1180",
            "1200", "1210", "1215", "1220", "1225", "1230", "1235", "1301",
            "1351", "1380", "1390", "1006", "1036", "1360", "1365", "1400"
        ];
        return includes(consultingMatters, _serviceArea.matterCode);
    };

    //Functionality:  Sets matter year to 00
    //Trigger:  Called from other methods as needed
    self.SetMatterYear00 = function () {
        setValue("cxp_matteryeartext", "00");
    };

    //Functionality: Sets matter year to 00
    //Trigger: Called from other methods as needed
    self.SetMatterYearToTaxReturnYear = function () {
        var alertStrings = {
            text: "You must enter a tax return year to default the matter year."
        };
        var returnYear = getValue("cxp_care_taxreturnyear");

        if (!returnYear) {
            Xrm.Navigation.openAlertDialog(alertStrings).then(
                function () {
                    setFocus("cxp_care_taxreturnyear");
                }, function () {
                    setFocus("cxp_care_taxreturnyear");
                });

        } else if (returnYear.length === 2) {
            setValue("cxp_matteryeartext", returnYear);
        } else if (returnYear.length > 2) {
            setValue("cxp_matteryeartext", returnYear.slice(-2));
        }
    };

    //Functionality:Shows an alert if the field is no
    //Trigger: onChange of fields 'cxp_careadvisoryhasattestpartnerdeterminedind'
    self.AttestPartnerDeterminedIndependentOnChange = function () {
        var alertStrings = {
            text: "Obtain approval from the Attest Partner."
        };
        var no = 280410001;
        var value = getValue("cxp_careadvisoryhasattestpartnerdeterminedind");
        if (value === no) {
            Xrm.Navigation.openAlertDialog(alertStrings);
        }
    };

    //Functionality: Shows an alert if field is No
    //Trigger: onChange of field 'cxp_careadvisorynationalauditdeterminedindepe'
    self.AuditDeterminedIndependentOnChange = function () {
        var alertStrings = {
            text: "Obtain approval from National Audit."
        };
        var no = 280410001;
        var value = getValue("cxp_careadvisorynationalauditdeterminedindepe");
        if (value === no) {
            Xrm.Navigation.openAlertDialog(alertStrings);
        }
    };

    //Functionality: Shows an alert if field is true
    //Trigger: onChange of field 'cxp_care_3sec'
    self.SecAdvisoryOnChange = function () {
        var alertStrings = {
            text: "Consult with Attest Partner and National Audit."
        };
        var value = getValue("cxp_care_3sec");
        if (value === true) {
            Xrm.Navigation.openAlertDialog(alertStrings);
        }
    };


    //Functionality: Shows an alert if field is true
    //Trigger:onChange of field 'cxp_care_3pcaob'
    self.PSCAOBAdvisoryOnChange = function () {
        var alertStrings = {
            text: "Consult with Attest Partner and National Audit."
        };
        var value = getValue("cxp_care_3pcaob");
        if (value === true) {
            Xrm.Navigation.openAlertDialog(alertStrings);
        }
    };

    //Functionality:Shows an alert if field is No
    //Trigger:onChange of field 'cxp_care_consideredtheeconomicfirmstandards'
    self.EconomonicFirmStandardsOnChange = function () {
        var alertStrings = {
            text: "You must consider the economic firm standard before submitting this request."
        };
        var no = 280410001;
        var value = getValue("cxp_care_consideredtheeconomicfirmstandards");
        if (value === no) {
            Xrm.Navigation.openAlertDialog(alertStrings);
        }
    };

    //Functionality: Throw alert at the user if service line doesn't have multiple service lines checked
    //Trigger: onChange of matterYear
    self.matterYearOnChange = function () {
        var serviceRequested = getValue("cxp_servicearea");
        var serviceRequestedID = serviceRequested ? serviceRequested[0].id : "";
        var matterYear = getValue("cxp_matteryeartext");
        var client = getValue("cxp_account");
        var clientId = client ? client[0].id : "";
        var opp = getValue("cxp_opportunity");
        var oppId = opp ? opp[0].id : "";
        serviceRequestedID = serviceRequestedID.replace("{", "").replace("}", "");
        clientId = clientId.replace("{", "").replace("}", "");
        oppId = oppId.replace("{", "").replace("}", "");
        if ((matterYear !== null) && (oppId !== "")) {
            var filter = "_cxp_product_value eq '" + serviceRequestedID + "' and _cpdc_clientid_value eq '" + clientId + "' and cxp_matteryeartext eq '" + matterYear + "' and _cxp_opportunityid_value ne '" + oppId + "' and statecode eq 0";
            var resultMatters = JSON.parse(getMultiple("cpdc_matters", "cpdc_name", filter, false));
            var status = getValue("statecode");
            if (status === 0) {
                if (_serviceArea.multipleServices === false) {
                    if (resultMatters.value.length >= 1) {
                        alert("You cannot use this service line as it permits only one service in an year. Select a new Service Line");
                        //setValue("cxp_servicerequested", null);
                    }
                }
                else {
                    if (resultMatters.value.length >= 5) {
                        alert("You cannot use this service line as it permits only one service in an year. Select a new Service Line");
                        //setValue("cxp_servicerequested", null);
                    }
                }
            }
        }
    };

    //Functionality:  Checks values to see if the 'Request Approvals' button should be visible.
    //Trigger:  Enable Rule for "Request Approvals" ribbon button
    self.shouldRequestApprovalsButtonBeVisible = function (executionContext) {
        var context = executionContext.getFormContext();
		var allowMovingToFinalizePhase = getValue("cxp_allowmovingtofinalizephase");
		var expectedMargin = getValue("cxp_expectedprofitability");

		//if (allowMovingToFinalizePhase) {
		if ((allowMovingToFinalizePhase) && (expectedMargin !== null)) {
			return true;
			
        } else {
            return false;
        }
    };

    //Functionality:  Checks values to see if the 'close as won' button should be visible.
    //Trigger:  Enable Rule for Close As Won ribbon button
    self.shouldCloseAsWonButtonBeVisible = function (executionContext) {
        //var wasImported = self.oppWasPreviouslyClosedFromImport();
        //var isFinalizeStage = self.BPFInFinalizeStage();
        var context = executionContext.getFormContext();
        var isFinalizeStage = context.getAttribute("cxp_phase").getValue() === 5;
        var allowMovingToFinalizePhase = getValue("cxp_allowmovingtofinalizephase");
        var isRCMSOpportunity = getValue("cxp_servicecategory")[0].name === "RCMS";

        //if ((self.isOpportunityApproved() || self.isRCMSOpportunity()) &&
        if ((self.isOpportunityApproved() || isRCMSOpportunity) &&
            isFinalizeStage && allowMovingToFinalizePhase && !self.isClosedAsWon() && !self.isClosedAsLost()) {
            return true;
        } else {
            return false;
        }
    };

    //Functionality: If the RCMS section is showing then the opportunity Rik extension is of type RCMS.
    self.isRCMSOpportunity = function () {
        var rcmsSection = getSection(getTab("RCMS"), "RCMS");
        return getSectionVisible(rcmsSection);
    };

    //Goes through all of the fields and makes sure that they should be approved.
    self.isOpportunityApproved = function () {

        var possibleFilled = 0;
        var actualFilled = 0;

        approvalPairs.forEach(function (pair) {
            var approval = pair[0];
            var signature = pair[1];

            if (getVisible(approval) || getVisible(signature)) {
                possibleFilled += 1;
                if (getValue(approval) && getValue(signature)) {
                    actualFilled += 1;
                }
            }
        });

        return possibleFilled !== 0 && actualFilled !== 0
            && possibleFilled === actualFilled;
    };

    // verify if already closed as won
    self.isClosedAsWon = function () {
        var closedAsWon = getValue("cxp_closedaswon");
        if (closedAsWon)
            return true;
        else
            return false;
    };

    // verify if already closed as lost
    self.isClosedAsLost = function () {
        var closedAsLost = getValue("cxp_closedaslost");
        if (closedAsLost)
            return true;
        else
            return false;
    };

    //Functionality: Checks that all of the approvals that need filled in are completely filled in.
    //  If they are filled in then the opportunity approved field is filled set to true.
    //Triggers: 
    //  cxp_engagementletterbypassapproval, cxp_engagementletterbypasssignature, cxp_partnerprincipalapprover
    //  cxp_partnerprincipalsignature, cxp_ompapprover, cxp_ompsignature
    //  cxp_riskreviewerapprover, cxp_riskreviewersignature, cxp_taxpracticedirectorapprover
    //  cxp_taxpracticedirectorsignature, cxp_nationalleaderapprover, cxp_nationalleadersignature
    //  cxp_nationalleadersignature, cxp_cannabistaxdirectorsignature"], cxp_cannabistaxpartnerapprover
    //  cxp_cannabistaxpartnersignature, cxp_professionalpracticeleaderapprover, cxp_professionalpracticeleaderbypasssignature,
    //  cxp_engagementpartner, cxp_engagementpartnersignature
  //  self.CheckApprovals = function () {
  //      var possibleFilled = 0;
  //      var actualFilled = 0;

  //      approvalPairs.forEach(function (pair) {
  //          var approval = pair[0];
  //          var signature = pair[1];

  //          if (getVisible(approval) || getVisible(signature)) {
  //              possibleFilled += 1;
  //              if (getValue(approval) && getValue(signature)) {
  //                  actualFilled += 1;
  //              }
  //          }
  //      });
		////Xrm.Navigation.openAlertDialog("test:" + possibleFilled + "test" + actualFilled);
  //      if ((possibleFilled !== 0) && (actualFilled !== 0) && (possibleFilled === actualFilled)) {
  //          setValue("cxp_opportunityapproved", true, true);
  //      } else {
  //          setValue("cxp_opportunityapproved", false, true);
		//}
		////context.data.entity.save();
  //      refreshRibbon(true);
  //  };

	self.CheckApprovals = function () {
		var oppApproved = self.isOpportunityApproved();
		if (oppApproved) {
			setValue("cxp_opportunityapproved", true, true);
		}
		else {
			setValue("cxp_opportunityapproved", false, true);
		}
		context.data.entity.save();
	};


    //Functionality: Throws dialouge box at the user when "Create matter" flag is sets to No
    //Trigger: On click of Close As Won(Custom) button
    self.closeAsWonAlertDialouge = function (executionContext) {
        if (getValue("cxp_creatematter") === false) {
            var confirmStrings = { text: "Do you want to close this opportunity as won without creating a matter?", title: "Matter Creation", confirmButtonLabel: "Yes", cancelButtonLabel: "No" };
            var confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    if (success.confirmed) {
                        setValue("cxp_creatematter", false);
                        self.winOpp(executionContext.getFormContext());
                        //setValue("cxp_closedaswon", true);
                        console.log("Dialog closed using OK button.");
                    } else {
                        setValue("cxp_creatematter", true);
                        self.winOpp(executionContext.getFormContext());
                        //setValue("cxp_closedaswon", true);
                        console.log("Dialog closed using Cancel button or X.");
                    }
                });
        } else {
            self.winOpp(executionContext.getFormContext());
            //setValue("cxp_closedaswon", true);
        }
    };

    //Functionality: Closes the OpportunityRiskExtension  as won status
    //Trigger: This is being called from self.closeAsWonAlertDialouge
    self.winOpp = function (formContext) {
        // check for acct number or Next lient Number
        var account = getValue("cxp_account");
        var accountId = account ? account[0].id : "";
        var resultACC = JSON.parse(getSingle("accounts", accountId, "cxp_nextclientnumber ,accountnumber", false));
        var nextClientNumber = resultACC["cxp_nextclientnumber"];
        var accountNumber = resultACC["accountnumber"];

        if (nextClientNumber === null && accountNumber === null) {
            setValue("cxp_closedaswon", false, true);
            formContext.data.entity.save();
            var alertStrings2 = {
                text: "Please submit a Sphinx ticket to get assistance from the CXP team  : Missing Account Number, Opportunity cannot be Closed as Won"
            };
            Xrm.Navigation.openAlertDialog(alertStrings2);
        }
        else {
            //     
            setValue("cxp_closedaswon", true, true);

            formContext.data.entity.save();
            var alertStrings = {
                text: "The opportunity is now closed as won."
            };
            Xrm.Navigation.openAlertDialog(alertStrings);

        }

        //    var id = formContext.data.entity.getId();
        //    id = id.replace("{", "").replace("}", "");
        //    var opportunityclose = {
        //        "opportunityid@odata.bind": "/opportunities(" + id + ")",
        //        "actualrevenue": 100,
        //        "actualend": new Date(),
        //        "description": "Your description here"
        //    };

        //    var parameters = {
        //        "OpportunityClose": opportunityclose,
        //        "Status": -1
        //    };
        //    var context;
        //    if (typeof getglobalcontext === "function") {
        //        context = getglobalcontext();
        //    } else {
        //        context = Xrm.Page.context;
        //    }

        //    var req = new XMLHttpRequest();
        //    req.open("POST", context.getClientUrl() + "/api/data/v8.2/WinOpportunity", true);
        //    req.setRequestHeader("OData-MaxVersion", "4.0");
        //    req.setRequestHeader("OData-Version", "4.0");
        //    req.setRequestHeader("Accept", "application/json");
        //    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        //    req.onreadystatechange = function () {
        //        if (this.readyState === 4) {
        //            req.onreadystatechange = null;
        //            if (this.status === 204) {
        //                Xrm.Page.data.refresh(true).then("", "");
        //            } else {
        //                var errorText = this.responseText;
        //                //Error and errorText variable contains an error - do something with it
        //            }
        //        }
        //    };
        //    req.send(JSON.stringify(parameters));
    };

    self.setDefaultValueofBillingContact = function () {
        var account = getValue("cxp_account");
        if (account !== null && account !== undefined && account.length > 0) {
            var accountID = account[0].id;
            var filter = "_cxp_contactid_value eq '" + accountID + "' and cxp_primarycontact eq true";
            if (contactResult !== undefined)
                var contactResult = JSON.parse(getMultiple("contacts", "fullname", filter, false));
            if (contactResult !== null && contactResult !== undefined && contactResult.value.length > 0) {
                var contact = contactResult.value[0];
                var contactName = contact.fullname;
                var contactID = contact.contactid;
                setLookupValue("cxp_billingcontact", contactID, contactName, "contact", null);
            }
        }
    };


    self.lockFieldsIfClosedAsWon = function () {
        var closedAsWon = getValue("cxp_closedaswon");
        var closedAsLost = getValue("cxp_ClosedAsLost");
        var controls = context.ui.controls.get();
        if (closedAsWon || closedAsLost) {
            for (var i in controls) {
                var ctrl = controls[i];
                var ctrlType = ctrl.getControlType();
                if (ctrl && ctrlType !== "subgrid" && ctrlType !== "iframe" && ctrlType !== "webresource") {

                    ctrl.setDisabled(true);
                }
            }
        }

    };

    //

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
            //setValue("cxp_elbypassreason", null);
            setRequiredLevel("cxp_elbypassreason", "required");
        }
        else {
            setVisible("cxp_elbypassreason", false);
            setValue("cxp_elbypassreason", null);
            setRequiredLevel("cxp_elbypassreason", "none");
        }
    };



    //Billing and Engagement attached fields need to be completed before request for Approval and close as won
    self.requiredBeforeCloseAsWon = function (executionContext) {
        var context = executionContext.getFormContext();
        var allowMovingToFinalizePhase = getValue("cxp_allowmovingtofinalizephase");

        if (allowMovingToFinalizePhase) {
            setRequiredLevel("cxp_billingstreet1", "required");
            //setRequiredLevel("cxp_billingcontact", "required");
            setRequiredLevel("cxp_apclerk", "required");
            setRequiredLevel("cxp_billingstateprovince", "required");
            setRequiredLevel("cxp_billingcity", "required");
            setRequiredLevel("cxp_billingzippostalcode", "required");
            setRequiredLevel("cxp_engagementletterattached", "required");
            setRequiredLevel("cxp_matterlocation", "required");



        }
        else {
            setRequiredLevel("cxp_billingstreet1", "none");
            setRequiredLevel("cxp_billingcontact", "none");
            setRequiredLevel("cxp_apclerk", "none");
            setRequiredLevel("cxp_billingstateprovince", "none");
            setRequiredLevel("cxp_billingcity", "none");
            setRequiredLevel("cxp_billingzippostalcode", "none");
            setRequiredLevel("cxp_engagementletterattached", "none");
            setRequiredLevel("cxp_matterlocation", "required");
        }
    };

    //show hide Matter Details Subgrid
	/*self.showMatterDetailsGrid = function () {
		 
		var matterName = getValue("cxp_matter");

		if (matterName !== null  ) {
	   setSectionVisibleFromTab(getTab("FinalizeTab"), "Matters Details", true)
		}
        else
         {
         setSectionVisibleFromTab(getTab("FinalizeTab"), "Matters Details", false)
         }
	};

*/

    self.setPhaseto6 = function () {
        var phase = getValue("cxp_phase");
        setValue("cxp_phase", 5);
    };

    //  
    self.setDefaultViewMatterLocation = function () {
        setLookupViewByName("cxp_matterlocation", "GL Offices", true);
    };

    // Default view on Lookup field 


    self.setDefaultViewBTK = function () {
        setLookupViewByName("cxp_billingtimekeeper", "All CXP Users", true);
    };

    self.setDefaultViewOTK = function () {
        setLookupViewByName("cxp_originatingtimekeeper", "All CXP Users", true);
    };


    self.setDefaultViewSTK = function () {
        setLookupViewByName("cxp_supervisingtimekeeper", "All CXP Users", true);
    };



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

    //copy opp product and service area from Opp to RE on Load of RE form
    self.copyProductAndServicAreaOnLoad = function (formContext) {
		var matterRE = getValue("cxp_matter");

        var opp = getValue("cxp_opportunity");
        var oppId = opp ? opp[0].id : "";
        oppId = oppId.replace("{", "").replace("}", "");
        //var oppProductRE = getValue("cxp_opportunityproduct");  //cxp_opportunityproduct
        // var serviceAreaRE = getValue("cxp_servicecategory") ;  //cxp_servicecategory
        var resultOpp = JSON.parse(getSingle("opportunities", oppId, "_cxp_opportunityproduct_value , _pricelevelid_value,estimatedvalue", false));
        var oppProduct = resultOpp["_cxp_opportunityproduct_value"];
		var servArea = resultOpp["_pricelevelid_value"];
		var expectedFee = resultOpp["estimatedvalue"];  
        var _cxp_opportunityproduct_value_formatted = resultOpp["_cxp_opportunityproduct_value@OData.Community.Display.V1.FormattedValue"];
        //var _cxp_opportunityproduct_value_lookuplogicalname = resultOpp["_cxp_opportunityproduct_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
        var _pricelevelid_value_formatted = resultOpp["_pricelevelid_value@OData.Community.Display.V1.FormattedValue"];
        //var _pricelevelid_value_lookuplogicalname = resultOpp["_pricelevelid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

		setValue("cxp_expectedfees", expectedFee); // updatimg expected fee on RE for calculating expected margin $

		if (oppProduct !== null && oppProduct !== "" && servArea !== null && servArea !== "") {
            //setLookupValue("cxp_opportunityproduct", oppProduct.Id, _cxp_opportunityproduct_value_formatted, "OpportunityProduct", null);
			setLookupValue("cxp_opportunityproduct", oppProduct, _cxp_opportunityproduct_value_formatted, "OpportunityProduct", null);
			//setLookupValue("cxp_servicecategory", servArea.Id, _pricelevelid_value_formatted, "pricelevel", null);
			setLookupValue("cxp_servicecategory", servArea, _pricelevelid_value_formatted, "pricelevel", null);
		}
		context.data.entity.save();
    };

    /*
    // to set the default view of lookup of subgrid with the view “attest service areas” 
    var attestAreasViewID = 'A638911B-F7F0-E911-A995-000D3A1D5F25';  //A638911B-F7F0-E911-A995-000D3A1D5F25

    //      

    self.setDefaultViewAttestAreas = function (formContext) {
        debugger;
        var attestAreasViewID = 'A638911B-F7F0-E911-A995-000D3A1D5F25';
                     
       var SubgridLookUp = context.ui.controls.get("AdvisoryAttestServicesWePerform");
                //var SubgridLookUp = getControl("AdvisoryAttestServicesWePerform");
        if (SubgridLookUp != null && SubgridLookUp != undefined)
        {
                    SubgridLookUp.setDefaultView("{" + attestAreasViewID + "}");
                    SubgridLookUp.refresh();
                    
        }
          
    }
    */
	

    //#endregion
    return {
        onLoad: self.onLoad,
        onSave: self.onSave,
        shouldCloseAsWonButtonBeVisible: self.shouldCloseAsWonButtonBeVisible,
        shouldRequestApprovalsButtonBeVisible: self.shouldRequestApprovalsButtonBeVisible,
        closeAsWonAlertDialouge: self.closeAsWonAlertDialouge


    };

})();
