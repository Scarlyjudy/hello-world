"use strict";
var FORM_CXP_OPPORTUNITY = "67312889-1a99-4553-9af1-11f3a3fa3027";
var FORM_CXP_SHORT_FORM = "559ed811-b776-4036-b8e7-727a2e6e3681";
var FORM_TAX_SHORT_FORM = "9645bd5a-0850-477c-93d3-0721987e0c16";

//Phone Number Formatting JavaScript
function FormatPhoneNumber(executionContext, fieldName) {
    var formContext = executionContext.getFormContext();
    var oField = formContext.getAttribute(fieldName);
	var sAllNumeric = oField;
	//Xrm.Navigation.openAlertDialog(alertStrings, 
	var sFormattedPhoneNumber;
	var validated;
    if (typeof (oField) !== "undefined" && oField.getValue() !== null && oField.getValue().toLowerCase().indexOf("x") < 0) {
        sAllNumeric = oField.getValue().replace(/[^0-9]/g, "");
        switch (sAllNumeric.length) {
			case 10:
				validated = validateNumber(sAllNumeric);
				if (validated) {
					sFormattedPhoneNumber = "(" + sAllNumeric.substr(0, 3) + ") " + sAllNumeric.substr(3, 3) + "-" + sAllNumeric.substr(6, 4);
					oField.setValue(sFormattedPhoneNumber);
				}
				else {
					oField.setValue(null);
				}
                break;
			case 11:
				validated = validateNumber(sAllNumeric);
				if (validated) {
					sFormattedPhoneNumber = "+" + sAllNumeric.substr(0, 1) + " (" + sAllNumeric.substr(1, 3) + ") " + sAllNumeric.substr(4, 3) + "-" + sAllNumeric.substr(7, 4);
					oField.setValue(sFormattedPhoneNumber);
				}
				else {
					oField.setValue(null);
				}
                break;
			case 12:
				validated = validateNumber(sAllNumeric);
				if (validated) {
					sFormattedPhoneNumber = "+" + sAllNumeric.substr(0, 2) + " (" + sAllNumeric.substr(2, 3) + ") " + sAllNumeric.substr(5, 3) + "-" + sAllNumeric.substr(8, 4);
					oField.setValue(sFormattedPhoneNumber);
				}
				else {
					oField.setValue(null);
				}
                break;
			case 14:
				validated = validateNumber(sAllNumeric);
				if (validated) {
					sFormattedPhoneNumber = "+" + sAllNumeric.substr(0, 1) + "-" + sAllNumeric.substr(1, 3) + " (" + sAllNumeric.substr(4, 3) + ") " + sAllNumeric.substr(7, 3) + "-" + sAllNumeric.substr(10, 4);
					oField.setValue(sFormattedPhoneNumber);
				}
				else {
					oField.setValue(null);
				}
                break;
            default:
                //alert("Phone numbers must contain 10, 11, 12 or 14 digits.  For example: 111-111-1111 or 1 111-111-1111 or 11 111-111-111 or 1 111 111-111-1111.")
				var alertStrings = { confirmButtonLabel: "Ok", text: "Phone numbers must contain 10, 11, 12 or 14 digits. For example: XXX-XXX-XXXX or X XXX-XXX-XXXX or XX XXX-XXX-XXXX or X XXX XXX-XXX-XXXX." };
				var alertOptions = { height: 180, width: 350 };
				Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                 break;
        }
    }
}

function validateNumber(phNum) {
	if ((phNum === "1234567890") || (phNum === "0000000000") || (phNum === "1111111111") || (phNum === "2222222222") || (phNum === "3333333333") || (phNum === "4444444444") || (phNum === "5555555555") || (phNum === "6666666666") || (phNum === "7777777777") || (phNum === "8888888888") || (phNum === "9999999999")
		|| (phNum === "00000000000") || (phNum === "11111111111") || (phNum === "22222222222") || (phNum === "33333333333") || (phNum === "44444444444") || (phNum === "55555555555") || (phNum === "66666666666") || (phNum === "77777777777") || (phNum === "88888888888") || (phNum === "99999999999")
		|| (phNum === "000000000000") || (phNum === "111111111111") || (phNum === "222222222222") || (phNum === "333333333333") || (phNum === "444444444444") || (phNum === "555555555555") || (phNum === "666666666666") || (phNum === "777777777777") || (phNum === "888888888888") || (phNum === "999999999999")
		|| (phNum === "00000000000000") || (phNum === "11111111111111") || (phNum === "22222222222222") || (phNum === "33333333333333") || (phNum === "44444444444444") || (phNum === "55555555555555") || (phNum === "66666666666666") || (phNum === "77777777777777") || (phNum === "88888888888888") || (phNum === "99999999999999")) {
		var alertStrings = { confirmButtonLabel: "Ok", text: "Please enter a valid Phone Number." };
		var alertOptions = { height: 180, width: 350 };
		Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
		return false;
	}
	else {
		return true;
	}
}

function FormatEIN(executionContext, fieldName) {
    var formContext = executionContext.getFormContext();
    var oField = formContext.getAttribute(fieldName);
    var sAllNumeric = oField;
    var sFormattedPhoneNumber;
    if (typeof (oField) !== "undefined" && oField.getValue() !== null) {
        sAllNumeric = oField.getValue().replace(/[^0-9]/g, "");
        sFormattedPhoneNumber = sAllNumeric.substr(0, 2) + "-" + sAllNumeric.substr(2, sAllNumeric.length - 1);
        oField.setValue(sFormattedPhoneNumber);
    }
}

function retrieveEntityById(entityType, entityId, query) {
    entityId = entityId.replace("{", "").replace("}", "");
    var globalContext = Xrm.Utility.getGlobalContext();
    var entity = null;
    var clientUrl = globalContext.getClientUrl();
    var webAPIPath = "/api/data/v8.1";
    var entityUri = "/" + entityType;
    var uri = clientUrl + webAPIPath + entityUri + "(" + entityId + ")?" + query;
    var request = new XMLHttpRequest();
    request.open("GET", encodeURI(uri), false);
    request.setRequestHeader("OData-MaxVersion", "4.0");
    request.setRequestHeader("OData-Version", "4.0");
    request.setRequestHeader("Accept", "application/json");
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    request.setRequestHeader("Prefer", 'odata.include-annotations="*"');
    request.onreadystatechange = function () {
        if (this.readyState === 4) {
            request.onreadystatechange = null;
            switch (this.status) {
                case 200: // Success with content returned in response body.
                case 204: // Success with no content returned in response body.
                    entity = JSON.parse(request.response);
                    break;
                default: // All other statuses are unexpected so are treated like errors.
                    var error;
                    try {
                        error = JSON.parse(request.response).error;
                    } catch (e) {
                        error = new Error("Unexpected Error");
                    }
                    //reject(new Error(error));
                    break;
            }
        }
    };
    request.send();

    return entity;

}

function retrieveMultiple(entityType, query) {
    //var formContext = executionContext.getFormContext();
    var globalContext = Xrm.Utility.getGlobalContext();
    var results = null;
    var clientUrl = globalContext.getClientUrl();
    var webAPIPath = "/api/data/v8.1";
    var entityUri = "/" + entityType;
    var uri = clientUrl + webAPIPath + entityUri + "?" + query;
    var request = new XMLHttpRequest();
    request.open("GET", encodeURI(uri), false);
    request.setRequestHeader("OData-MaxVersion", "4.0");
    request.setRequestHeader("OData-Version", "4.0");
    request.setRequestHeader("Accept", "application/json");
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    request.setRequestHeader("Prefer", 'odata.include-annotations="*"');
    request.onreadystatechange = function () {
        if (this.readyState === 4) {
            request.onreadystatechange = null;
            switch (this.status) {
                case 200: // Success with content returned in response body.
                case 204: // Success with no content returned in response body.
                    results = JSON.parse(request.response);
                    break;
                default: // All other statuses are unexpected so are treated like errors.
                    var error;
                    try {
                        error = JSON.parse(request.response).error;
                    } catch (e) {
                        error = new Error("Unexpected Error");
                    }
                    //reject(new Error(error));
                    break;
            }
        }
    };
    request.send();

    return results;

}

function createRecord(entityType, newObject) {
	
	var globalContext = Xrm.Utility.getGlobalContext();
	var clientUrl = globalContext.getClientUrl();
    var webAPIPath = "/api/data/v8.1";
    var entityUri = "/" + entityType;
    var uri = clientUrl + webAPIPath + entityUri;
    var request = new XMLHttpRequest();
    request.open("POST", encodeURI(uri), false);
    request.setRequestHeader("OData-MaxVersion", "4.0");
    request.setRequestHeader("OData-Version", "4.0");
    request.setRequestHeader("Accept", "application/json");
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    request.setRequestHeader("Prefer", 'odata.include-annotations="*"');
    request.onreadystatechange = function () {
        if (this.readyState === 4) {
            request.onreadystatechange = null;
            switch (this.status) {
                case 200: // Success with content returned in response body.
                case 204: // Success with no content returned in response body.
                    results = JSON.parse(request.response);
                    break;
                default: // All other statuses are unexpected so are treated like errors.
                    var error;
                    try {
                        error = JSON.parse(request.response).error;
                    } catch (e) {
                        error = new Error("Unexpected Error");
                    }
                    //reject(new Error(error));
                    break;
            }
        }
    };

    request.send(JSON.stringify(newObject));
}

function deleteRecord(recordType, recordId) {
	var globalContext = Xrm.Utility.getGlobalContext();
	var clientUrl = globalContext.getClientUrl();
    var webAPIPath = "/api/data/v8.1";
    var entityUri = "/" + recordType;
    var uri = clientUrl + webAPIPath + entityUri + "(" + recordId.replace("{", "").replace("}", "") + ")";
    var request = new XMLHttpRequest();
    request.open("DELETE", encodeURI(uri), false);
    request.setRequestHeader("OData-MaxVersion", "4.0");
    request.setRequestHeader("OData-Version", "4.0");
    request.setRequestHeader("Accept", "application/json");
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    request.setRequestHeader("Prefer", 'odata.include-annotations="*"');
    request.onreadystatechange = function () {
        if (this.readyState === 4) {
            request.onreadystatechange = null;
            switch (this.status) {
                case 200: // Success with content returned in response body.
                case 204: // Success with no content returned in response body.
                    results = JSON.parse(request.response);
                    break;
                default: // All other statuses are unexpected so are treated like errors.
                    var error;
                    try {
                        error = JSON.parse(request.response).error;
                    } catch (e) {
                        error = new Error("Unexpected Error");
                    }
                    //reject(new Error(error));
                    break;
            }
        }
    };

    request.send();
}

function GetStateOrProvince(stateId) {
    stateId = stateId.replace("{", "").replace("}", "");
    var state = retrieveEntityById("cpdc_states", stateId, "$select=_cxp_country_value");
    return state;
}

function GetUser(userId, columnSet) {
    userId = userId.replace("{", "").replace("}", "");
    var user = retrieveEntityById("systemusers", userId, columnSet);
    return user;
}

function GetCRIndustry(crIndustryId, columnSet) {
    crIndustryId = crIndustryId.replace("{", "").replace("}", "");
    var crIndustry = retrieveEntityById("cpdc_industrypracticeareas", crIndustryId, columnSet);
    return crIndustry;
}

function GetSecurityRoleIdsByName(name) {
    //var datakey = "CrmSecurityRoleIds_" + name;
    //if (typeof (Storage) !== "undefined") {
    //    // Get the data from cache using user-specific cache key
    //    var res = sessionStorage.getItem(datakey);

    //    if (!res) {
    //        var roles = retrieveMultiple("roles", "$select=roleid&$filter=name eq '" + name + "'");
    //        sessionStorage.setItem(datakey, roles);
    //        return roles;
    //    } else {
    //        return (res.indexOf("true") !== -1)
    //    }
    //} else {
    //    return retrieveMultiple("roles", "$select=roleid&$filter=name eq '" + name + "'");
    //}


    var roles = retrieveMultiple("roles", "$select=roleid&$filter=name eq '" + name + "'");
    //var role = roles.value[0];
    return roles;
}

function GetUserByName(name) {
    var users = retrieveMultiple("systemusers", "$select=systemuserid,fullname&$filter=fullname eq '" + name + "'");
    var user = users.value[0];
    return user;
}

function GetSystemParameterValue(name) {
    var returnValue = "";

    //var datakey = "CrmSystemParameter_" + name;
    //if (typeof (Storage) !== "undefined") {
    //    // Get the data from cache using user-specific cache key
    //    var res = localStorage.getItem(datakey);

    //    if (!res) {
    //        var parameters = retrieveMultiple("cxp_systemparameters", "$select=cxp_name,cxp_value&$filter=cxp_name eq '" + name + "'");
    //        if (parameters !== null && parameters.value[0] !== null) {
    //            var parameter = parameters.value[0];
    //            localStorage.setItem(datakey, parameter);
    //            returnValue = parameter["cxp_value"];
    //        }
    //    } else {
    //        return (res.indexOf("true") !== -1)
    //    }
    //} else {
    //    var parameters = retrieveMultiple("cxp_systemparameters", "$select=cxp_name,cxp_value&$filter=cxp_name eq '" + name + "'");
    //    if (parameters !== null && parameters.value[0] !== null) {
    //        var parameter = parameters.value[0];
    //        returnValue = parameter["cxp_value"];
    //    }
    //}


    var parameters = retrieveMultiple("cxp_systemparameters", "$select=cxp_name,cxp_value&$filter=cxp_name eq '" + name + "'");
    if (parameters !== null && parameters.value[0] !== null) {
        var parameter = parameters.value[0];
        returnValue = parameter["cxp_value"];
    }

    return returnValue;
}

function GetCountryByName(name) {
    var returnValue = null;
    var countries = retrieveMultiple("cpdc_countries", "$select=cpdc_countryid, cpdc_name&$filter=cpdc_name eq '" + name + "'");
    if (countries !== null && countries.value[0] !== null) {
        returnValue = countries.value[0];
    }

    return returnValue;
}


function GetBillingCoordinatorRoleIds() {
    var roleIds = null;
    //var role = GetSecurityRoleIdByName("CXP Billing Coordinator");
    //if (role !== null) {
    //roleId = role["roleid"];
    //}

    roleIds = GetSecurityRoleIdsByName("CXP - Billing Coordinator");

    return roleIds;
}

function IsUserSystemAdministrator() {
 
    var bReturnValue = false;
    var context = Xrm.Utility.getGlobalContext();
    var userRoles = context.userSettings.securityRoles
    var systemAdminRoleIds = GetSystemAdminRoleIds();

    for (var i = 0; i < userRoles.length; i++) {
        var roleId = userRoles[i];
        for (var j = 0; j < systemAdminRoleIds.value.length; j++) {
            if (roleId === systemAdminRoleIds.value[j]["roleid"]) {
                bReturnValue = true;
                break;
            }
        }

        if (bReturnValue === true) {
            break;
        }
    }

    return bReturnValue;
}

function GetSystemAdminRoleIds() {
    var roleIds = null;
    //var role = GetSecurityRoleIdByName("CXP Billing Coordinator");
    //if (role !== null) {
    //roleId = role["roleid"];
    //}

    roleIds = GetSecurityRoleIdsByName("System Administrator");

    return roleIds;
}

function GetServiceAreaServiceType(serviceAreaId) {
    var serviceType = "";
    serviceAreaId = serviceAreaId.replace("{", "").replace("}", "");
    var serviceArea = retrieveEntityById("rg_serviceareas", serviceAreaId, "$select=cpdc_servicetype");

    if (serviceArea !== null) {
        serviceType = serviceArea["cpdc_servicetype"];
    }

    return serviceType;
}

function GetOffice(officeId, columnSet) {
    var office = null;
    if (officeId !== null) {
        officeId = officeId.replace("{", "").replace("}", "");
        office = retrieveEntityById("rg_offices", officeId, columnSet);
    }

    return office;
}

function GetTaxPracticeDirectorForBTK(btkUser) {
    var officeId = btkUser["_cxp_gloffice_value"];
    var officeName = btkUser["_cxp_gloffice_value@OData.Community.Display.V1.FormattedValue"];

    var office = GetOffice(officeId, "$select=_cxp_region_value");
    var region = "";
    if (office !== null) {
        region = office["_cxp_region_value@OData.Community.Display.V1.FormattedValue"];
        if (region !== null && region !== "" && region === "National Tax") {
            officeName = "National";
        }
    }

    if (officeName.indexOf("&") >= 0) {
        officeName = officeName.replace("&", "");
    }

    if (officeName.indexOf("-") >= 0) {
        officeName = officeName.replace("-", "");
    }
    var taxPracticeDirector = GetSystemParameterValue(officeName + " Tax Practice Director 1");

    if (taxPracticeDirector !== null && taxPracticeDirector !== "") {
        var tpdUser = GetUserByName(taxPracticeDirector);
        var tpdRef = null;

        if (tpdUser !== null) {
            tpdRef = new Array();
            tpdRef[0] = new Object();
            tpdRef[0].id = tpdUser["systemuserid"];
            tpdRef[0].entityType = "systemuser";
            tpdRef[0].name = tpdUser["fullname"];
        }
    }

    return tpdRef;
}

function callAction(actionName, entityName, actionParams, formContext) {
	var globalContext = Xrm.Utility.getGlobalContext();
    var result = null;
	var oDataEndPoint = encodeURI(globalContext.getClientUrl() + "/api/data/v8.1/" + entityName + "(" + formContext.data.entity.getId().replace("{", "").replace("}", "") + ")/Microsoft.Dynamics.CRM." + actionName);

    var request = new XMLHttpRequest();
    request.open("POST", oDataEndPoint, false);
    request.setRequestHeader("OData-MaxVersion", "4.0");
    request.setRequestHeader("OData-Version", "4.0");
    request.setRequestHeader("Accept", "application/json");
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    //request.setRequestHeader("Prefer", 'odata.include-annotations="*"');
    request.onreadystatechange = function () {
        if (request.readyState === 4) {
            request.onreadystatechange = null;
            if (request.status === 200) {
                result = JSON.parse(this.response);
            } else {
                var error = JSON.parse(this.response).error;
                Xrm.Navigation.openAlertDialog("Error calling action: " + actionName + error.message);
            }
        }
    };

    if (actionParams !== null)
        request.send(JSON.stringify(actionParams));
    else
        request.send();

    return result;
}

function GetCXPFollow(recordId, userId) {
    var cxpFollowId = null;

    var query = "$select=activityid&$filter=_regardingobjectid_value eq " + recordId.replace("{", "").replace("}", "");
    query += " and _ownerid_value eq " + userId.replace("{", "").replace("}", "");
    var cxpFollows = retrieveMultiple("cxp_cxpfollows", query);

    if (cxpFollows !== null && cxpFollows.value.length > 0) {
        cxpFollowId = cxpFollows.value[0]["activityid"];
    }

    return cxpFollowId;
}


