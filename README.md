var _isFinalizeStage = false;

var oppGlobal = {};

function po_onLoad(executionContext) {
    po_addAllOnChange(executionContext.getFormContext());
    po_opportunityREApproval(executionContext);
    //executionContext.getFormContext().data.process.addOnStageChange(po_opportunityApproval);
}
function po_opportunityREApproval(executionContext) {
    var yes = 280410000;
    var assuranceShortWithEngagmentPartnerApprover = 280410006;
    var allowMovingToFinalizePhase = getValue("cxp_allowmovingtofinalizephase");
    var formContext = executionContext.getFormContext();
    //_isFinalizeStage = formContext.data.process.getActiveStage().getName() === "Finalize";
    _isFinalizeStage = formContext.getAttribute("cxp_phase").getValue() === 5;
    var opportunityRE = po_getSetOpportunityREData(formContext);
    var svcRequest = opportunityRE.cxp_servicerequested.getValue();
    var oppo = opportunityRE.cxp_opportunity.getValue();
    po_getOpp(oppo[0].id, formContext, opportunityRE);
    var formType = formContext.ui.getFormType();
    if ((formType !== 1 && formType !== 2 && formType !== 4) || svcRequest === null) {
        console.log("return");
        return;
    }
    //if (opportunity.cxp_islongform.getValue() === yes && _isFinalizeStage) {
	//if (opportunityRE.cxp_islongform.getValue() === yes && _isFinalizeStage && allowMovingToFinalizePhase) {
	if (opportunityRE.cxp_islongform.getValue() === yes && _isFinalizeStage) {
        po_getServiceArea(svcRequest[0].id, formContext, opportunityRE, oppGlobal);
    }
    else {
        if ((opportunityRE.cxp_formtype.getValue() === assuranceShortWithEngagmentPartnerApprover) && (_isFinalizeStage)) {
        //if ((opportunityRE.cxp_formtype.getValue() === assuranceShortWithEngagmentPartnerApprover) && (_isFinalizeStage) && (allowMovingToFinalizePhase)) {
            po_assuranceShortWithEngagmentPartnerApprover(opportunityRE, oppGlobal);
        }
        else if (allowMovingToFinalizePhase) {
            po_shortFormPath(opportunityRE, oppGlobal);
        }
	}
	if (allowMovingToFinalizePhase) {
		po_toggleApprovalAttributesOnChange(executionContext);
	}
}

function po_addAllOnChange(formContext) {
    po_addOnChange(formContext, "cxp_careadvisoryhasattestpartnerdeterminedind", po_opportunityREApproval);
    //po_addOnChange(formContext, "rg_servicingofficeid", po_opportunityREApproval);
    po_addOnChange(formContext, "cxp_carearethereanyconcernsaboutconflictofint", po_opportunityREApproval);
    po_addOnChange(formContext, "cxp_audit_risk", po_opportunityREApproval);
    //po_addOnChange(formContext, "cxp_isthisacannabisrelatedopportunity", po_opportunityREApproval);
    po_addOnChange(formContext, "cxp_careindicatetheclientsrisklevel", po_opportunityREApproval);
    //po_addOnChange(formContext, "estimatedvalue", po_opportunityREApproval);
    po_addOnChange(formContext, "cxp_departmentcode", po_opportunityREApproval);
    //po_addOnChange(formContext, "cxp_billingtimekeeper", po_opportunityREApproval);
    //po_addOnChange(formContext, "cxp_servicerequested", po_opportunityREApproval);
    //po_addOnChange(formContext, "cxp_clientnumber", po_opportunityREApproval);
    //po_addOnChange(formContext, "cxp_relatedclientnumber", po_opportunityREApproval);
    po_addOnChange(formContext, "cxp_islongform", po_opportunityREApproval);
}

function po_addOnChange(formContext, fieldName, onChangeCallBack) {
    var attribute = formContext.getAttribute(fieldName);

    if (attribute) {
        attribute.addOnChange(onChangeCallBack);
    } else {
        window.console.warn("Attempted to add an OnChange callback to field " +
            fieldName +
            " but the field was not on the form");
    }
}

function po_handleServiceType(svcAreaData, formContext, opportunityRE, oppGlobal) {
    console.log("***Enter po_handleServiceType***");
    var svcType = svcAreaData.servicetype.toLowerCase();

    switch (svcType) {
        case "tax compliance":
        case "tax consulting":
        case "tax consulting - international":
        case "tax consulting - scott h":
        case "tax consulting - salt":
        case "tax consulting - trust and estate":
            po_clearAllElsePathFields(opportunityRE);
            po_clearAssurancePathFields(opportunityRE);
            po_taxRelatedPath(opportunityRE, svcAreaData, oppGlobal);
            break;
        case "assurance":
            po_clearTaxPathFields(opportunityRE);
            po_clearAllElsePathFields(opportunityRE);
            po_assurancePath(opportunityRE, oppGlobal);
            break;
        case "rcms":
            po_clearTaxPathFields(opportunityRE);
            po_clearAllElsePathFields(opportunityRE);
            po_clearAssurancePathFields(opportunityRE);
            break;
        default:
            po_clearAssurancePathFields(opportunityRE);
            po_clearTaxPathFields(opportunityRE);
            po_allElsePath(opportunityRE, oppGlobal);
            break;
    }
    formContext.data.entity.save();
}

function po_handleIsNewServiceForFamily(oppGlobal, svcAreaData, opportunityRE) {
    console.log("***Enter po_handleIsNewServiceForFamily***");
    //var billingTimeKeeper = oppGlobal.billingTimeKeeper.getValue();
    var billingTimeKeeper = new Array();
    billingTimeKeeper[0] = new Object();
    billingTimeKeeper[0].id = oppGlobal.billingTimeKeeper;
    billingTimeKeeper[0].name = oppGlobal.billingTimeKeeper_name;
    billingTimeKeeper[0].entityType = oppGlobal.billingTimeKeeper_entity;
    console.log(svcAreaData.taxleaderapproval);
    console.log(svcAreaData.taxpracticedirectorapproval);
    console.log(svcAreaData.othertaxleaderapproval);

    if (svcAreaData.taxleaderapproval) {
        //po_getSetTaxLeader(opportunity);
        po_getSetTaxLeader(opportunityRE);
    }
    else if (svcAreaData.taxpracticedirectorapproval) {
        console.log("***Inside SvcAreaData.taxpracticedirectorapproval branch***");
        //console.log(billingTimeKeeper[0].name);
        if (billingTimeKeeper[0].name !== null && billingTimeKeeper[0].name !== undefined) {
            console.log("***Set Tax Practice Director Approver to Billing Time Keeper***");
            po_getSetUser(billingTimeKeeper, opportunityRE.cxp_taxpracticedirectorapprover, "cxp_taxpracticedirector", opportunityRE);
        }
    }
    else if (svcAreaData.othertaxleaderapproval) {
        console.log("***Set Tax Practice Director to User with 'other tax leader = yes'");
        //po_getSetOtherTaxLeader(opportunity);
        po_getSetOtherTaxLeader(opportunityRE);
    }
    else {
        console.log("**Set Tax Practice Director to NULL");
        if (opportunityRE.cxp_taxpracticedirectorapprover.getValue()) {
            opportunityRE.cxp_taxpracticedirectorapprover.setValue(null);
        }
    }
}

function po_shortFormPath(opportunityRE, oppGlobal) {
    po_clearTaxPathFields(opportunityRE);
    po_clearAssurancePathFields(opportunityRE);
    po_clearAllElsePathFields(opportunityRE);

    //var billingTimeKeeper = oppGlobal.billingTimeKeeper.getValue();
    var billingTimeKeeper = new Array();
    billingTimeKeeper[0] = new Object();
    billingTimeKeeper[0].id = oppGlobal.billingTimeKeeper;
    billingTimeKeeper[0].name = oppGlobal.billingTimeKeeper_name;
    billingTimeKeeper[0].entityType = oppGlobal.billingTimeKeeper_entity;
    if (billingTimeKeeper[0].name !== null && billingTimeKeeper[0].name !== undefined) {
        //alert("if btk contains data");
        setLookup(opportunityRE.partnerPrincipalApprover, billingTimeKeeper[0].id, billingTimeKeeper[0].name, billingTimeKeeper[0].entityType);
    }
}

function po_assuranceShortWithEngagmentPartnerApprover(opportunityRE, oppGlobal) {
    po_clearTaxPathFields(opportunityRE);
    po_clearAssurancePathFields(opportunityRE);
    po_clearAllElsePathFields(opportunityRE);

    var pursuitLead = oppGlobal.pursuitLead.getValue();
    if (pursuitLead !== null) {
        //alert("if btk contains data");
        setLookup(opportunityRE.assuranceEngagementApprover, pursuitLead[0].id, pursuitLead[0].name, pursuitLead[0].entityType);
    }
}

function po_taxRelatedPath(opportunityRE, svcAreaData, oppGlobal) {
    console.log("***enter po_taxRelatedPath***");
    //var billingTimeKeeper = opportunity.billingTimeKeeper.getValue();
    //var billingTimeKeeper = opp.billingTimeKeeper.getValue();
    var billingTimeKeeper = new Array();
    billingTimeKeeper[0] = new Object();
    billingTimeKeeper[0].id = oppGlobal.billingTimeKeeper;
    billingTimeKeeper[0].name = oppGlobal.billingTimeKeeper_name;
    billingTimeKeeper[0].entityType = oppGlobal.billingTimeKeeper_entity;
    var departmentCode = opportunityRE.cxp_departmentcode.getValue();

    if (billingTimeKeeper[0].name !== null && billingTimeKeeper[0].name !== undefined) {
        //alert("if btk contains data");
        setLookup(opportunityRE.partnerPrincipalApprover, billingTimeKeeper[0].id, billingTimeKeeper[0].name, billingTimeKeeper[0].entityType);
    }

    //alert(opportunity.cxp_departmentcode.getValue());
    if (departmentCode !== null && departmentCode === 280410001) { //Tax
        //alert("inside tax if");
        console.log("***set OMPApprover = null***");
        if (opportunityRE.ompApprover.getValue()) {
            opportunityRE.ompApprover.setValue(null);
        }
    }

    if (oppGlobal.cxp_clientnumber === null && oppGlobal.cxp_relatedclientnumber === null) {
        po_handleIsNewServiceForFamily(oppGlobal, svcAreaData, opportunityRE);
    }
    else if (oppGlobal.cxp_relatedclientnumber !== null && svcAreaData.mattercode === null) {
        po_getIsNewServiceForFamily(svcAreaData, oppGlobal, opportunityRE);
    }

    if (oppGlobal.cxp_engagementletterattached === 280410001) { //No - Exception Requested
        po_getSetTaxEngagementLetterBypassApprover(opportunityRE);
    }

    if (oppGlobal.estimatedvalue > 100000) {
        po_getSetNationalTaxDirector(opportunityRE);
    }
    else {
        if (opportunityRE.cxp_nationalleaderapprover.getValue()) {
            opportunityRE.cxp_nationalleaderapprover.setValue(null);
        }
    }

	if (opportunityRE.cxp_careindicatetheclientsrisklevel.getValue() === 280410003) {//NA
		//cxp_departmentcode = TAX
		if (opportunityRE.cxp_departmentcode.getValue !== 280410001) {
			opportunityRE.cxp_departmentcode.setValue(280410001); //TAX
		}
		if (opportunityRE.cxp_careindicatetheclientsrisklevel.getValue()) {
			opportunityRE.cxp_careindicatetheclientsrisklevel.setValue(null);
		}
	}


    if (oppGlobal.cxp_isthisacannabisrelatedopportunity === 280410001) { //Yes
        po_getSetCannabisTaxPartner(opportunityRE);
        po_getSetCannabisTaxPracticeDirector(opportunityRE);
    } else {
        opportunityRE.cxp_cannabistaxpartnerapprover.setValue(null);
        opportunityRE.cxp_cannabistaxdirectorapprover.setValue(null);
    }

    if (svcAreaData.taxleaderapproval) {
        po_getSetTaxLeader(opportunityRE);
    }
    else if (svcAreaData.taxpracticedirectorapproval) {
        console.log("***Inside SvcAreaData.taxpracticedirectorapproval branch***");

        if (billingTimeKeeper[0].name !== null && billingTimeKeeper[0].name !== undefined) {
            console.log("***Set Tax Practice Director Approver to Billing Time Keeper***");
            po_getSetUser(billingTimeKeeper, opportunityRE.cxp_taxpracticedirectorapprover, "cxp_taxpracticedirector", opportunityRE);
        }
    }
    else if (svcAreaData.othertaxleaderapproval) {
        console.log("***Set Tax Practice Director to User with 'other tax leader = yes'");
        po_getSetOtherTaxLeader(opportunityRE);
    }
    else {
        console.log("**Set Tax Practice Director to NULL");
        if (opportunityRE.cxp_taxpracticedirectorapprover.getValue()) {
            opportunityRE.cxp_taxpracticedirectorapprover.setValue(null);
        }
    }
}

function po_assurancePath(opportunityRE, oppGlobal) {
    console.log("***Entering Assurance Path***");
    //var billingTimeKeeper = oppGlobal.billingTimeKeeper.getValue();
    var billingTimeKeeper = new Array();
    billingTimeKeeper[0] = new Object();
    billingTimeKeeper[0].id = oppGlobal.billingTimeKeeper;
    billingTimeKeeper[0].name = oppGlobal.billingTimeKeeper_name;
    billingTimeKeeper[0].entityType = oppGlobal.billingTimeKeeper_entity;
    //var office = oppGlobal.rg_servicingofficeid.getValue();
    var office = oppGlobal.rg_servicingofficeid;
    if (billingTimeKeeper[0].name !== null && billingTimeKeeper[0].name !== undefined) {
        setLookup(opportunityRE.partnerPrincipalApprover, billingTimeKeeper[0].id, billingTimeKeeper[0].name, billingTimeKeeper[0].entityType);

        po_getSetUser(billingTimeKeeper, opportunityRE.ompApprover, "cxp_officemanagingpartner", opportunityRE);
    }

    if (oppGlobal.cxp_engagementletterattached === 280410001) { //No - Exception Requested

        po_getSetAssuranceLeader(opportunityRE.cxp_engagementletterbypassapprover);
    }

    if (!(getValue("cxp_careassurancetierlevel") === 280410002)) { //not Tier III, for all other tiers, get National leader Approval

        po_getSetAssuranceLeader(opportunityRE.cxp_nationalleaderapprover);
    }

    //po_getSetAssuranceLeader(opportunityRE.cxp_nationalleaderapprover);  //******Remove for Kevin hedden, added in above filter
    if (opportunityRE.cxp_departmentcode.getValue() !== 280410000) {
        opportunityRE.cxp_departmentcode.setValue(280410000); // = AUD;
    }

    if (opportunityRE.cxp_audit_risk.getValue() === 280410003) {//NA
        if (opportunityRE.cxp_audit_risk.getValue()) {
            opportunityRE.cxp_audit_risk.setValue(null);
        }
    }

    if (opportunityRE.cxp_carearethereanyconcernsaboutconflictofint.getValue() === 280410000) {//Yes
        po_getSetRiskLeader(opportunityRE);
    }

    if (office !== null) {
        //po_getSetOffice(office[0].id, opportunityRE.cxp_professionalpracticeleaderapprover, "cxp_professionalpracticeleader", opportunityRE);
        po_getSetOffice(office, opportunityRE.cxp_professionalpracticeleaderapprover, "cxp_professionalpracticeleader", opportunityRE);
    }
}

function po_allElsePath(opportunityRE, oppGlobal) {
    //var billingTimeKeeper = oppGlobal.billingTimeKeeper.getValue();
    var billingTimeKeeper = new Array();
    billingTimeKeeper[0] = new Object();
    billingTimeKeeper[0].id = oppGlobal.billingTimeKeeper;
    billingTimeKeeper[0].name = oppGlobal.billingTimeKeeper_name;
    billingTimeKeeper[0].entityType = oppGlobal.billingTimeKeeper_entity;
    var expectedProfitability = oppGlobal.cxp_expectedprofitability;

    if (billingTimeKeeper[0].name !== null && billingTimeKeeper[0].name !== undefined) {
        setLookup(opportunityRE.partnerPrincipalApprover, billingTimeKeeper[0].id, billingTimeKeeper[0].name, billingTimeKeeper[0].entityType);

        po_getSetUser(billingTimeKeeper, opportunityRE.ompApprover, "cxp_officemanagingpartner", opportunityRE);
    }

    if (oppGlobal.cxp_engagementletterattached === 280410001) { //No - Exception Requested

        po_getSetAdvisoryLeader(opportunityRE.cxp_engagementletterbypassapprover);
    }


    if (oppGlobal.estimatedvalue > 200000 || (expectedProfitability < 25 && expectedProfitability)) {

        po_getSetAdvisoryLeader(opportunityRE.cxp_nationalleaderapprover);
    }
    else {
        if (opportunityRE.cxp_nationalleaderapprover.getValue()) {
            opportunityRE.cxp_nationalleaderapprover.setValue(null);
        }
    }

    if (opportunityRE.cxp_departmentcode.getValue()) {
        opportunityRE.cxp_departmentcode.setValue(280410002); //CON
    }

    if (opportunityRE.cxp_careadvisorynationalauditdeterminedindepe.getValue() === 280410001) {//No
        //THEN – cxp_ riskreviewerapprover = ANY USER IN CRM with cxp_ riskleader = YES
        po_getSetRiskLeader(opportunityRE);
    }
}

function po_getServiceArea(svcArea, formContext, opportunityRE, oppGlobal) {
    console.log("***Enter po_getServiceArea***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/products(" + standardizeGUID(svcArea) + ")?$select=cxp_mattercode,cxp_servicetype,cxp_othertaxleaderapproval,cxp_taxleaderapproval,cxp_taxpracticedirectorapproval",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var result = data;
            var svcAreaData = {};
            svcAreaData.mattercode = result["cxp_mattercode"];
            svcAreaData.servicetype = result["cxp_servicetype"];
            svcAreaData.othertaxleaderapproval = result["cxp_othertaxleaderapproval"];
            svcAreaData.taxleaderapproval = result["cxp_taxleaderapproval"];
            svcAreaData.taxpracticedirectorapproval = result["cxp_taxpracticedirectorapproval"];
            po_handleServiceType(svcAreaData, formContext, opportunityRE, oppGlobal);
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getOpp(oppId, formContext, opportunityRE) {
    console.log("***Enter po_getOpp***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        //url: globalContext.getClientUrl() + "/api/data/v8.2/opportunities(" + standardizeGUID(oppId) + ")?$select=ownerid,cxp_billingtimekeeper,cxp_clientnumber,cxp_relatedclientnumber,cxp_engagementletterattached, estimatedvalue, cxp_isthisacannabisrelatedopportunity, rg_servicingofficeid, estimatedvalue, cxp_expectedprofitability",
        url: globalContext.getClientUrl() + "/api/data/v8.2/opportunities(" + standardizeGUID(oppId) + ")?$select=_ownerid_value,_cxp_billingtimekeeper_value,cxp_clientnumber,cxp_relatedclientnumber,cxp_engagementletterattached,cxp_isthisacannabisrelatedopportunity,_rg_servicingofficeid_value, estimatedvalue, cxp_expectedprofitability",
        //url: globalContext.getClientUrl() + "/api/data/v8.2/opportunities(" + standardizeGUID(oppId) + ")?$select=_ownerid_value",

        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var result = data;
            //var opp = {};
            oppGlobal.pursuitLead = result["_ownerid_value"];
            oppGlobal.billingTimeKeeper = result["_cxp_billingtimekeeper_value"];
            oppGlobal.billingTimeKeeper_entity = result["_cxp_billingtimekeeper_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
            oppGlobal.billingTimeKeeper_name = result["_cxp_billingtimekeeper_value@OData.Community.Display.V1.FormattedValue"];
            oppGlobal.cxp_clientnumber = result["cxp_clientnumber"];
            oppGlobal.cxp_relatedclientnumber = result["cxp_relatedclientnumber"];
            oppGlobal.cxp_engagementletterattached = result["cxp_engagementletterattached"];
            oppGlobal.estimatedvalue = result["estimatedvalue"];
            oppGlobal.cxp_isthisacannabisrelatedopportunity = result["cxp_isthisacannabisrelatedopportunity"];
            oppGlobal.rg_servicingofficeid = result["_rg_servicingofficeid_value"];
            //opp.estimatedvalue = result["estimatedvalue"];
            oppGlobal.cxp_expectedprofitability = result["cxp_expectedprofitability"];
            //oppGlobal = opp;
            //return opp;

            //po_handleServiceType(svcAreaData, formContext, opportunityRE);
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetOffice(office, attributeToSet, officeAttributeToGet, opportunityRE) {
    console.log("***Enter po_getSetOffice***");
    var globalContext = Xrm.Utility.getGlobalContext();

    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/rg_offices(" + standardizeGUID(office) + ")",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var taxPracticeDirector = [];
            var officeManagingPartner = [];
            var professionalPracticeLeader = [];
            var result = data;
            var rg_officeid = result["rg_officeid"];
            var cxp_professionalpracticeleader_fullname = result["_cxp_professionalpracticeleader_value@OData.Community.Display.V1.FormattedValue"];
            var cxp_professionalpracticeleader_systemuserid = result["_cxp_professionalpracticeleader_value"];
            professionalPracticeLeader.push({ "id": cxp_professionalpracticeleader_systemuserid, "name": cxp_professionalpracticeleader_fullname });
            console.log(professionalPracticeLeader);
            var cxp_officemanagingpartner_fullname = result["_cxp_officemanagingpartner_value@OData.Community.Display.V1.FormattedValue"];
            var cxp_officemanagingpartner_systemuserid = result["_cxp_officemanagingpartner_value"];
            officeManagingPartner.push({ "id": cxp_officemanagingpartner_systemuserid, "name": cxp_officemanagingpartner_fullname });
            console.log(officeManagingPartner);
            var cxp_taxpracticedirector_fullname = result["_cxp_taxpracticedirector_value@OData.Community.Display.V1.FormattedValue"];
            var cxp_taxpracticedirector_systemuserid = result["_cxp_taxpracticedirector_value"];
            taxPracticeDirector.push({ "id": cxp_taxpracticedirector_systemuserid, "name": cxp_taxpracticedirector_fullname });
            console.log(taxPracticeDirector);

            if (officeAttributeToGet === "cxp_officemanagingpartner") {
                po_getSetUser(officeManagingPartner, attributeToSet, null, opportunityRE);
            }
            if (officeAttributeToGet === "cxp_professionalpracticeleader") {
                po_getSetUser(professionalPracticeLeader, attributeToSet, null, opportunityRE);
            }
            if (officeAttributeToGet === "cxp_taxpracticedirector") {
                if (cxp_taxpracticedirector_systemuserid !== null) {
                    po_getSetUser(taxPracticeDirector, attributeToSet, null, opportunityRE);
                }
                else {
                    //po_getSetTaxLeader(opportunity);
                    po_getSetTaxLeader(opportunityRE);
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetUser(user/*Array*/, attributeToSet, officeAttributeToGet, opportunityRE) {//set officeAttribute to null when bypassing the getOffice path
    console.log("***Enter po_getSetUser***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers(" + standardizeGUID(user[0].id) + ")?$select=cxp_advisoryleader,cxp_assuranceleader,cxp_billingtimekeeper,cxp_cannabistaxpartner,cxp_cannabistaxpracticedirector,_cxp_gloffice_value,cxp_nationaltaxdirector,cxp_riskleader,cxp_taxengagementletterbypassapprover,cxp_taxleader,fullname",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var result = data;
            var userData = {};
            userData.cxp_advisoryleader = result["cxp_advisoryleader"];
            userData.cxp_assuranceleader = result["cxp_assuranceleader"];
            userData.cxp_billingtimekeeper = result["cxp_billingtimekeeper"];
            userData.cxp_cannabistaxpartner = result["cxp_cannabistaxpartner"];
            userData.cxp_cannabistaxpracticedirector = result["cxp_cannabistaxpracticedirector"];
            userData.cxp_gloffice = result["_cxp_gloffice_value"];
            userData.cxp_nationaltaxdirector = result["cxp_nationaltaxdirector"];
            userData.cxp_riskleader = result["cxp_riskleader"];
            userData.cxp_taxengagementletterbypassapprover = result["cxp_taxengagementletterbypassapprover"];
            userData.cxp_taxleader = result["cxp_taxleader"];
            userData.fullname = result["fullname"];

            if (officeAttributeToGet === null) {
                setLookup(attributeToSet, user[0].id, user[0].name, "systemuser");
            }
            else {
                if (userData.cxp_gloffice !== null) {
                    po_getSetOffice(userData.cxp_gloffice, attributeToSet, officeAttributeToGet, opportunityRE);
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getIsNewServiceForFamily(svcAreaData, oppGlobal, opportunityRE) {
    console.log("***Enter po_getIsNewServiceForFamily***");
    var relatedClientNumber = oppGlobal.cxp_relatedclientnumber.getValue();
    var matterCode = svcAreaData.matterCode;

    if (matterCode === null || relatedClientNumber === null || opportunityRE.opportunityId === null) {
        return false;
    }
    var parameters = {};
    parameters.RelatedClientNumber = relatedClientNumber;
    parameters.MatterCode = matterCode;

    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/opportunities(" + standardizeGUID(opportunityRE.opportunityId) + ")/Microsoft.Dynamics.CRM.cxp_IsNewServiceForClientFamily",
        data: JSON.stringify(parameters),
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results.NewService) {
                po_handleIsNewServiceForFamily(oppGlobal, svcAreaData, opportunityRE);
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetAdvisoryLeader(attributeToSet) {
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_advisoryleader eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var advisoryLeader = {};
                advisoryLeader.fullname = results.value[0]["fullname"];
                advisoryLeader.systemuserid = results.value[0]["systemuserid"];

                if (advisoryLeader.systemuserid !== null) {
                    setLookup(attributeToSet, advisoryLeader.systemuserid, advisoryLeader.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetRiskLeader(opportunityRE) {
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_riskleader eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var riskLeader = {};
                riskLeader.fullname = results.value[0]["fullname"];
                riskLeader.systemuserid = results.value[0]["systemuserid"];

                if (riskLeader.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_riskreviewerapprover, riskLeader.systemuserid, riskLeader.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetCannabisTaxPartner(opportunityRE) {
    console.log("***Enter po_getSetCannabisTaxPartner***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_cannabistaxpartner eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var cannabisTaxPartner = {};
                cannabisTaxPartner.fullname = results.value[0]["fullname"];
                cannabisTaxPartner.systemuserid = results.value[0]["systemuserid"];

                if (cannabisTaxPartner.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_cannabistaxpartnerapprover, cannabisTaxPartner.systemuserid, cannabisTaxPartner.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetCannabisTaxPracticeDirector(opportunityRE) {
    console.log("***Enter po_getSetCannabisTaxPracticeDirector***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_cannabistaxpracticedirector eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var cannabisTaxPracticeDirector = {};
                cannabisTaxPracticeDirector.fullname = results.value[0]["fullname"];
                cannabisTaxPracticeDirector.systemuserid = results.value[0]["systemuserid"];

                if (cannabisTaxPracticeDirector.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_cannabistaxdirectorapprover, cannabisTaxPracticeDirector.systemuserid, cannabisTaxPracticeDirector.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetTaxEngagementLetterBypassApprover(opportunityRE) {
    console.log("***Enter po_getSetTaxEngagementLetterBypassApprover***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_taxengagementletterbypassapprover eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var taxEngagementLetterBypassApprover = {};
                taxEngagementLetterBypassApprover.fullname = results.value[0]["fullname"];
                taxEngagementLetterBypassApprover.systemuserid = results.value[0]["systemuserid"];

                if (taxEngagementLetterBypassApprover.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_engagementletterbypassapprover, taxEngagementLetterBypassApprover.systemuserid, taxEngagementLetterBypassApprover.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetTaxLeader(opportunityRE) {
    console.log("***Enter po_getSetTaxLeader***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_taxleader eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var taxLeader = {};
                taxLeader.fullname = results.value[0]["fullname"];
                taxLeader.systemuserid = results.value[0]["systemuserid"];

                if (taxLeader.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_taxpracticedirectorapprover, taxLeader.systemuserid, taxLeader.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetNationalTaxDirector(opportunityRE) {
    console.log("***Enter po_getSetNationalTaxDirector***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_nationaltaxdirector eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var nationalTaxDirector = {};
                nationalTaxDirector.fullname = results.value[0]["fullname"];
                nationalTaxDirector.systemuserid = results.value[0]["systemuserid"];

                if (nationalTaxDirector.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_nationalleaderapprover, nationalTaxDirector.systemuserid, nationalTaxDirector.fullname, "systemuser");
                }

            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetOtherTaxLeader(opportunityRE) {
    console.log("***Enter po_getSetOtherTaxLeader***");
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_othertaxleader eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var otherTaxLeader = {};
                otherTaxLeader.fullname = results.value[0]["fullname"];
                otherTaxLeader.systemuserid = results.value[0]["systemuserid"];

                if (otherTaxLeader.systemuserid !== null) {
                    setLookup(opportunityRE.cxp_taxpracticedirectorapprover, otherTaxLeader.systemuserid, otherTaxLeader.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetAssuranceLeader(attributeToSet) {
    var globalContext = Xrm.Utility.getGlobalContext();
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: globalContext.getClientUrl() + "/api/data/v8.2/systemusers?$select=fullname,systemuserid&$filter=cxp_assuranceleader eq true",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        },
        async: !_isFinalizeStage,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results !== null) {
                var assuranceLeader = {};
                assuranceLeader.fullname = results.value[0]["fullname"];
                assuranceLeader.systemuserid = results.value[0]["systemuserid"];

                if (assuranceLeader.systemuserid !== null) {
                    setLookup(attributeToSet, assuranceLeader.systemuserid, assuranceLeader.fullname, "systemuser");
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Navigation.openAlertDialog(textStatus + " " + errorThrown);
        }
    });
}

function po_getSetOpportunityREData(formContext) {
    var opportunityREData = {};
    opportunityREData.cxp_islongform = formContext.getAttribute("cxp_islongform");
    opportunityREData.cxp_shortformreason = formContext.getControl("cxp_shortformreason");
    opportunityREData.cxp_formtype = formContext.getAttribute("cxp_formtype");
    //portunityData.cxp_servicerequested = formContext.getAttribute("cxp_servicerequested");
    opportunityREData.cxp_servicerequested = formContext.getAttribute("cxp_servicearea");
    opportunityREData.cxp_opportunity = formContext.getAttribute("cxp_opportunity");

    opportunityREData.cxp_departmentcode = formContext.getAttribute("cxp_departmentcode");
    opportunityREData.cxp_careindicatetheclientsrisklevel = formContext.getAttribute("cxp_careindicatetheclientsrisklevel");
    opportunityREData.cxp_audit_risk = formContext.getAttribute("cxp_audit_risk");
    opportunityREData.cxp_carearethereanyconcernsaboutconflictofint = formContext.getAttribute("cxp_carearethereanyconcernsaboutconflictofint");
    opportunityREData.cxp_careadvisorynationalauditdeterminedindepe = formContext.getAttribute("cxp_careadvisorynationalauditdeterminedindepe");


    //pportunityData.billingTimeKeeper = formContext.getAttribute("cxp_billingtimekeeper");
    //opportunityData.pursuitLead = formContext.getAttribute("ownerid");
    opportunityREData.assuranceEngagementApprover = formContext.getAttribute("cxp_engagementpartner"); // same as pursuit lead for assurance short with engagment partner as approver
    opportunityREData.partnerPrincipalApprover = formContext.getAttribute("cxp_partnerprincipalapprover");
    opportunityREData.ompApprover = formContext.getAttribute("cxp_ompapprover");
    //opportunityData.cxp_engagementletterattached = formContext.getAttribute("cxp_engagementletterattached");
    opportunityREData.cxp_engagementletterbypassapprover = formContext.getAttribute("cxp_engagementletterbypassapprover");
    //opportunityData.estimatedvalue = formContext.getAttribute("estimatedvalue");
    //opportunityData.cxp_expectedprofitability = formContext.getAttribute("cxp_expectedprofitability");
    opportunityREData.cxp_nationalleaderapprover = formContext.getAttribute("cxp_nationalleaderapprover");

    //opportunityData.cxp_carearethereanyconcernsaboutconflictofint = formContext.getAttribute("cxp_carearethereanyconcernsaboutconflictofint");
    opportunityREData.cxp_riskreviewerapprover = formContext.getAttribute("cxp_riskreviewerapprover");
    opportunityREData.cxp_careadvisoryhasattestpartnerdeterminedind = formContext.getAttribute("cxp_careadvisoryhasattestpartnerdeterminedind");
    //opportunityData.cxp_careadvisorynationalauditdeterminedindepe = formContext.getAttribute("cxp_careadvisorynationalauditdeterminedindepe");
    //opportunityData.cxp_audit_risk = formContext.getAttribute("cxp_audit_risk");
    //opportunityData.rg_servicingofficeid = formContext.getAttribute("rg_servicingofficeid");
    opportunityREData.cxp_professionalpracticeleaderapprover = formContext.getAttribute("cxp_professionalpracticeleaderapprover");
    opportunityREData.opportunityId = formContext.getAttribute("cxp_opportunity");
    //opportunityData.cxp_clientnumber = formContext.getAttribute("cxp_clientnumber");
    //opportunityData.cxp_relatedclientnumber = formContext.getAttribute("cxp_relatedclientnumber");
    opportunityREData.cxp_taxpracticedirectorapprover = formContext.getAttribute("cxp_taxpracticedirectorapprover");

    opportunityREData.cxp_cannabistaxpartnerapprover = formContext.getAttribute("cxp_cannabistaxpartnerapprover");
    opportunityREData.cxp_cannabistaxdirectorapprover = formContext.getAttribute("cxp_cannabistaxdirectorapprover");
    //opportunityData.cxp_isthisacannabisrelatedopportunity = formContext.getAttribute("cxp_isthisacannabisrelatedopportunity");
    opportunityREData.cxp_engagementletterbypasssignature = formContext.getAttribute("cxp_engagementletterbypasssignature");
    opportunityREData.cxp_engagementletterbypassapprovaldate = formContext.getAttribute("cxp_engagementletterbypassapprovaldate");
    opportunityREData.cxp_engagementletterapprovedby = formContext.getAttribute("cxp_engagementletterapprovedby");
    opportunityREData.cxp_engagementpartnersignature = formContext.getAttribute("cxp_engagementpartnersignature");  //added new
    opportunityREData.cxp_engagementpartnerapprovaldate = formContext.getAttribute("cxp_engagementpartnerapprovaldate"); //added new
    opportunityREData.cxp_engagmentpartnerapprovedby = formContext.getAttribute("cxp_engagmentpartnerapprovedby"); //added new
    opportunityREData.cxp_partnerprincipallsignature = formContext.getAttribute("cxp_partnerprincipalsignature");
    opportunityREData.cxp_partnerprincipalapprovaldate = formContext.getAttribute("cxp_partnerprincipalapprovaldate");
    opportunityREData.cxp_partnerprincipalapprovedby = formContext.getAttribute("cxp_partnerprincipalapprovedby");
    opportunityREData.cxp_ompsignature = formContext.getAttribute("cxp_ompsignature");
    opportunityREData.cxp_ompapprovaldate = formContext.getAttribute("cxp_ompapprovaldate");
    opportunityREData.cxp_ompapprovedby = formContext.getAttribute("cxp_ompapprovedby");
    opportunityREData.cxp_riskreviewsignature = formContext.getAttribute("cxp_riskreviewersignature");
    opportunityREData.cxp_riskreviewerapprovaldate = formContext.getAttribute("cxp_riskreviewerapprovaldate");
    opportunityREData.cxp_riskreviewerapprovedby = formContext.getAttribute("cxp_riskreviewerapprovedby");
    opportunityREData.cxp_taxpracticedirectorsignature = formContext.getAttribute("cxp_taxpracticedirectorsignature");
    opportunityREData.cxp_taxpracticedirectorapprovaldate = formContext.getAttribute("cxp_taxpracticedirectorapprovaldate");
    opportunityREData.cxp_taxpracticedirectorapprovedby = formContext.getAttribute("cxp_taxpracticedirectorapprovedby");
    opportunityREData.cxp_nationalleadersignature = formContext.getAttribute("cxp_nationalleadersignature");
    opportunityREData.cxp_nationalleaderapprovaldate = formContext.getAttribute("cxp_nationalleaderapprovaldate");
    opportunityREData.cxp_nationalleaderapprovedby = formContext.getAttribute("cxp_nationalleaderapprovedby");
    opportunityREData.cxp_cannabistaxdirectorsignature = formContext.getAttribute("cxp_cannabistaxdirectorsignature");
    opportunityREData.cxp_cannabistaxdirectorapprovaldate = formContext.getAttribute("cxp_cannabistaxdirectorapprovaldate");
    opportunityREData.cxp_cannabistaxdirectorapprovedby = formContext.getAttribute("cxp_cannabistaxdirectorapprovedby");
    opportunityREData.cxp_cannabistaxpartnersignature = formContext.getAttribute("cxp_cannabistaxpartnersignature");
    opportunityREData.cxp_cannabistaxpartnerapprovaldate = formContext.getAttribute("cxp_cannabistaxpartnerapprovaldate");
    opportunityREData.cxp_cannabistaxpartnerapprovedby = formContext.getAttribute("cxp_cannabistaxpartnerapprovedby");
    opportunityREData.cxp_professionalpracticeleadersignature = formContext.getAttribute("cxp_professionalpracticeleadersignature");
    opportunityREData.cxp_professionalpracticeleaderapprovaldate = formContext.getAttribute("cxp_professionalpracticeleaderapprovaldate");
    opportunityREData.cxp_professionalpracticeleaderapprovedby = formContext.getAttribute("cxp_professionalpracticeleaderapprovedby");
    //opportunityData.cxp_parentopportunity = formContext.getAttribute("cxp_parentopportunity");

    opportunityREData.engagementLetterBypassObj = {};
    opportunityREData.engagementLetterBypassObj.approverAttribute = opportunityREData.cxp_engagementletterbypassapprover;
    opportunityREData.engagementLetterBypassObj.signatureAttribute = opportunityREData.cxp_engagementletterbypasssignature;
    opportunityREData.engagementLetterBypassObj.attributesToHide = [
        opportunityREData.cxp_engagementletterbypasssignature,
        opportunityREData.cxp_engagementletterbypassapprovaldate,
        opportunityREData.cxp_engagementletterapprovedby,
        opportunityREData.cxp_engagementletterbypassapprover
    ];

    opportunityREData.assuranceEngagementObj = {};
    opportunityREData.assuranceEngagementObj.approverAttribute = opportunityREData.assuranceEngagementApprover;
    opportunityREData.assuranceEngagementObj.signatureAttribute = opportunityREData.cxp_engagementpartnersignature;
    opportunityREData.assuranceEngagementObj.attributesToHide = [
        opportunityREData.cxp_engagementpartnersignature,
        opportunityREData.cxp_engagementpartnerapprovaldate,
        opportunityREData.cxp_engagmentpartnerapprovedby,
        opportunityREData.assuranceEngagementApprover
    ];

    opportunityREData.partnerPrincipalObj = {};
    opportunityREData.partnerPrincipalObj.approverAttribute = opportunityREData.partnerPrincipalApprover;
    opportunityREData.partnerPrincipalObj.signatureAttribute = opportunityREData.cxp_partnerprincipallsignature;
    opportunityREData.partnerPrincipalObj.attributesToHide = [
        opportunityREData.cxp_partnerprincipallsignature,
        opportunityREData.cxp_partnerprincipalapprovaldate,
        opportunityREData.cxp_partnerprincipalapprovedby,
        opportunityREData.partnerPrincipalApprover
    ];

    opportunityREData.ompApproverObj = {}
    opportunityREData.ompApproverObj.approverAttribute = opportunityREData.ompApprover;
    opportunityREData.ompApproverObj.signatureAttribute = opportunityREData.cxp_ompsignature;
    opportunityREData.ompApproverObj.attributesToHide = [
        opportunityREData.cxp_ompsignature,
        opportunityREData.cxp_ompapprovaldate,
        opportunityREData.cxp_ompapprovedby,
        opportunityREData.ompApprover
    ];

    opportunityREData.riskReviewObj = {}
    opportunityREData.riskReviewObj.approverAttribute = opportunityREData.cxp_riskreviewerapprover;
    opportunityREData.riskReviewObj.signatureAttribute = opportunityREData.cxp_riskreviewsignature;
    opportunityREData.riskReviewObj.attributesToHide = [
        opportunityREData.cxp_riskreviewsignature,
        opportunityREData.cxp_riskreviewerapprovaldate,
        opportunityREData.cxp_riskreviewerapprovedby,
        opportunityREData.cxp_riskreviewerapprover
    ];

    opportunityREData.taxPracticeObj = {}
    opportunityREData.taxPracticeObj.approverAttribute = opportunityREData.cxp_taxpracticedirectorapprover;
    opportunityREData.taxPracticeObj.signatureAttribute = opportunityREData.cxp_taxpracticedirectorsignature;
    opportunityREData.taxPracticeObj.attributesToHide = [
        opportunityREData.cxp_taxpracticedirectorsignature,
        opportunityREData.cxp_taxpracticedirectorapprovaldate,
        opportunityREData.cxp_taxpracticedirectorapprovedby,
        opportunityREData.cxp_taxpracticedirectorapprover
    ];

    opportunityREData.nationalLeaderObj = {}
    opportunityREData.nationalLeaderObj.approverAttribute = opportunityREData.cxp_nationalleaderapprover;
    opportunityREData.nationalLeaderObj.signatureAttribute = opportunityREData.cxp_nationalleadersignature;
    opportunityREData.nationalLeaderObj.attributesToHide = [
        opportunityREData.cxp_nationalleadersignature,
        opportunityREData.cxp_nationalleaderapprovaldate,
        opportunityREData.cxp_nationalleaderapprovedby,
        opportunityREData.cxp_nationalleaderapprover
    ];

    opportunityREData.cannabisTaxDirectorObj = {}
    opportunityREData.cannabisTaxDirectorObj.approverAttribute = opportunityREData.cxp_cannabistaxdirectorapprover;
    opportunityREData.cannabisTaxDirectorObj.signatureAttribute = opportunityREData.cxp_cannabistaxdirectorsignature;
    opportunityREData.cannabisTaxDirectorObj.attributesToHide = [
        opportunityREData.cxp_cannabistaxdirectorsignature,
        opportunityREData.cxp_cannabistaxdirectorapprovaldate,
        opportunityREData.cxp_cannabistaxdirectorapprovedby,
        opportunityREData.cxp_cannabistaxdirectorapprover
    ];

    opportunityREData.cannabisTaxPartnerObj = {}
    opportunityREData.cannabisTaxPartnerObj.approverAttribute = opportunityREData.cxp_cannabistaxpartnerapprover;
    opportunityREData.cannabisTaxPartnerObj.signatureAttribute = opportunityREData.cxp_cannabistaxpartnersignature;
    opportunityREData.cannabisTaxPartnerObj.attributesToHide = [
        opportunityREData.cxp_cannabistaxpartnersignature,
        opportunityREData.cxp_cannabistaxpartnerapprovaldate,
        opportunityREData.cxp_cannabistaxpartnerapprovedby,
        opportunityREData.cxp_cannabistaxpartnerapprover
    ];

    opportunityREData.professionalPracticeObj = {}
    opportunityREData.professionalPracticeObj.approverAttribute = opportunityREData.cxp_professionalpracticeleaderapprover;
    opportunityREData.professionalPracticeObj.signatureAttribute = opportunityREData.cxp_professionalpracticeleadersignature;
    opportunityREData.professionalPracticeObj.attributesToHide = [
        opportunityREData.cxp_professionalpracticeleadersignature,
        opportunityREData.cxp_professionalpracticeleaderapprovaldate,
        opportunityREData.cxp_professionalpracticeleaderapprovedby,
        opportunityREData.cxp_professionalpracticeleaderapprover
    ];

    return opportunityREData;
}

function setLookup(attribute, attributeId, attributeName, attributeEntityType) {
    console.log("***Enter setLookup***");
    attribute.setValue([{ id: attributeId, name: attributeName, entityType: attributeEntityType }]);
    attribute.fireOnChange();
}

function standardizeGUID(id) {
    id = id.replace(/[{}]/g, "");
    return id;
}

function po_clearTaxPathFields(opportunityRE) {
    opportunityRE.partnerPrincipalApprover.setValue(null);
    opportunityRE.cxp_taxpracticedirectorapprover.setValue(null);
    opportunityRE.cxp_engagementletterbypassapprover.setValue(null);
    opportunityRE.cxp_nationalleaderapprover.setValue(null);
    //opportunity.cxp_departmentcode.setValue(null);
    opportunityRE.cxp_cannabistaxpartnerapprover.setValue(null);
    opportunityRE.cxp_cannabistaxdirectorapprover.setValue(null);
}

function po_clearAssurancePathFields(opportunityRE) {
    opportunityRE.partnerPrincipalApprover.setValue(null);
    opportunityRE.ompApprover.setValue(null);
    opportunityRE.cxp_engagementletterbypassapprover.setValue(null);
    opportunityRE.cxp_nationalleaderapprover.setValue(null);
    //opportunity.cxp_departmentcode.setValue(null);
    opportunityRE.cxp_riskreviewerapprover.setValue(null);
    opportunityRE.cxp_professionalpracticeleaderapprover.setValue(null);
}

function po_clearAllElsePathFields(opportunityRE) {
    opportunityRE.ompApprover.setValue(null);
    opportunityRE.cxp_engagementletterbypassapprover.setValue(null);
    opportunityRE.partnerPrincipalApprover.setValue(null);
    opportunityRE.cxp_nationalleaderapprover.setValue(null);
    //opportunity.cxp_departmentcode.setValue(null);
    opportunityRE.cxp_riskreviewerapprover.setValue(null);
    opportunityRE.assuranceEngagementApprover.setValue(null);
}

function po_toggleApprovalAttributes(formContext, approverAttribute, attributesToHide) { //onLoad, onChange
    console.log("Entering po_toggleApprovalAttributes");
    if (approverAttribute !== null && _isFinalizeStage) {
        for (var i = 0; i < attributesToHide.length; i++) {
            po_toggleAttributeVisibility(formContext, attributesToHide[i].getName(), true);
        }
    } else {
        for (var i = 0; i < attributesToHide.length; i++) {
            po_toggleAttributeVisibility(formContext, attributesToHide[i].getName(), false);
        }
    }
}

function po_toggleAttributeVisibility(formContext, attribute, isVisible) {
    if (formContext.getControl(attribute) !== null) {
        formContext.getControl(attribute).setVisible(isVisible);
    }
}

function po_setFormNotification(formContext) {
    var message = "One or more signatures are required before you will be able to close this Opportunity as won.";
    var level = "WARNING";
    var uniqueId = "1";
    //var opportunity = po_getSetOpportunityData(formContext);
    //var parentOpportunity = opportunity.cxp_parentopportunity.getValue();
    //if (_isFinalizeStage && parentOpportunity !== null && parentOpportunity !== undefined) {
    if (_isFinalizeStage) {
        formContext.ui.setFormNotification(message, level, uniqueId);
    }
}

function po_clearFormNotification(formContext) {
    formContext.ui.clearFormNotification("1");
}

function po_approvalSatisfied(approvalAttribute, sigAttribute) {
    var isSatisfied;

    if (approvalAttribute === null) {
        isSatisfied = true;
        return isSatisfied;
    }

    if (sigAttribute === null) {
        isSatisfied = false;
        return isSatisfied;
    }
    else {
        isSatisfied = true;
        return isSatisfied;
    }
}

function po_toggleApprovalAttributesOnChange(executionContext) {
    console.log("Entering po_toggleApprovalAttributesOnChange");
    var formContext = executionContext.getFormContext();
    var opportunityRE = po_getSetOpportunityREData(formContext);
    var engagementApprover = opportunityRE.engagementLetterBypassObj.approverAttribute;
    var engagementAttributesToHide = opportunityRE.engagementLetterBypassObj.attributesToHide;
    var assuranceEngagementApprover = opportunityRE.assuranceEngagementObj.approverAttribute;
    var assuranceEngagementAttributesToHide = opportunityRE.assuranceEngagementObj.attributesToHide;
    var partnerPrincipal = opportunityRE.partnerPrincipalObj.approverAttribute;
    var partnerPrincipalAttributesToHide = opportunityRE.partnerPrincipalObj.attributesToHide;
    var ompApprover = opportunityRE.ompApproverObj.approverAttribute;
    var ompApproverAttributesToHide = opportunityRE.ompApproverObj.attributesToHide;
    var riskReviewApprover = opportunityRE.riskReviewObj.approverAttribute;
    var riskReviewApproverAttributesToHide = opportunityRE.riskReviewObj.attributesToHide;
    var taxPracticeDirector = opportunityRE.taxPracticeObj.approverAttribute;
    var taxPracticeDirectorAttributesToHide = opportunityRE.taxPracticeObj.attributesToHide;
    var nationalLeader = opportunityRE.nationalLeaderObj.approverAttribute;
    var nationalLeaderAttributesToHide = opportunityRE.nationalLeaderObj.attributesToHide;
    var cannabisTaxDirector = opportunityRE.cannabisTaxDirectorObj.approverAttribute;
    var cannabisTaxDirectorAttributesToHide = opportunityRE.cannabisTaxDirectorObj.attributesToHide;
    var cannabisTaxPartner = opportunityRE.cannabisTaxPartnerObj.approverAttribute;
    var cannabisTaxPartnerAttributesToHide = opportunityRE.cannabisTaxPartnerObj.attributesToHide;
    var professionalPractice = opportunityRE.professionalPracticeObj.approverAttribute;
    var professionalPracticeAttributesToHide = opportunityRE.professionalPracticeObj.attributesToHide;

    if (engagementApprover !== null) {
        po_toggleApprovalAttributes(formContext, engagementApprover.getValue(), engagementAttributesToHide);
    }
    if (assuranceEngagementApprover !== null) {
        po_toggleApprovalAttributes(formContext, assuranceEngagementApprover.getValue(), assuranceEngagementAttributesToHide);
    }
    if (partnerPrincipal !== null) {
        po_toggleApprovalAttributes(formContext, partnerPrincipal.getValue(), partnerPrincipalAttributesToHide);
    }
    if (ompApprover !== null) {
        po_toggleApprovalAttributes(formContext, ompApprover.getValue(), ompApproverAttributesToHide);
    }
    if (riskReviewApprover !== null) {
        po_toggleApprovalAttributes(formContext, riskReviewApprover.getValue(), riskReviewApproverAttributesToHide);
    }
    if (taxPracticeDirector !== null) {
        po_toggleApprovalAttributes(formContext, taxPracticeDirector.getValue(), taxPracticeDirectorAttributesToHide);
    }
    if (nationalLeader !== null) {
        po_toggleApprovalAttributes(formContext, nationalLeader.getValue(), nationalLeaderAttributesToHide);
    }
    if (cannabisTaxDirector !== null) {
        po_toggleApprovalAttributes(formContext, cannabisTaxDirector.getValue(), cannabisTaxDirectorAttributesToHide);
    }
    if (cannabisTaxPartner !== null) {
        po_toggleApprovalAttributes(formContext, cannabisTaxPartner.getValue(), cannabisTaxPartnerAttributesToHide);
    }
    if (professionalPractice !== null) {
        po_toggleApprovalAttributes(formContext, professionalPractice.getValue(), professionalPracticeAttributesToHide);
    }

    po_handleFormNotification(opportunityRE, formContext);
    formContext.ui.refreshRibbon(true);
}

function po_handleFormNotification(opportunityRE, formContext) {
    console.log("Entering po_handleFormNotification");
    var isSatisfied = true;
    var opportunityREApprovers = [];
    if (opportunityRE.professionalPracticeObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.professionalPracticeObj.approverAttribute.getValue(), "signature": opportunityRE.professionalPracticeObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.cannabisTaxPartnerObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.cannabisTaxPartnerObj.approverAttribute.getValue(), "signature": opportunityRE.cannabisTaxPartnerObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.cannabisTaxDirectorObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.cannabisTaxDirectorObj.approverAttribute.getValue(), "signature": opportunityRE.cannabisTaxDirectorObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.nationalLeaderObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.nationalLeaderObj.approverAttribute.getValue(), "signature": opportunityRE.nationalLeaderObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.taxPracticeObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.taxPracticeObj.approverAttribute.getValue(), "signature": opportunityRE.taxPracticeObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.riskReviewObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.riskReviewObj.approverAttribute.getValue(), "signature": opportunityRE.riskReviewObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.ompApproverObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.ompApproverObj.approverAttribute.getValue(), "signature": opportunityRE.ompApproverObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.partnerPrincipalObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.partnerPrincipalObj.approverAttribute.getValue(), "signature": opportunityRE.partnerPrincipalObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.engagementLetterBypassObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.engagementLetterBypassObj.approverAttribute.getValue(), "signature": opportunityRE.engagementLetterBypassObj.signatureAttribute.getValue() });
    }
    if (opportunityRE.assuranceEngagementObj.approverAttribute !== null) {
        opportunityREApprovers.push({ "approval": opportunityRE.assuranceEngagementObj.approverAttribute.getValue(), "signature": opportunityRE.assuranceEngagementObj.signatureAttribute.getValue() });
    }
    // console.log(opportunityApprovers);

    opportunityREApprovers.forEach(function (element) {
        if (po_approvalSatisfied(element.approval, element.signature) === false) {
            po_setFormNotification(formContext);
            isSatisfied = false;
            return;
        }
    });

    if (isSatisfied) {
        po_clearFormNotification(formContext);
    }
}
