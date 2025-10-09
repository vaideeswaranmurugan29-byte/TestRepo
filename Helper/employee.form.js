"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var EmployeeForm;
(function (EmployeeForm) {
    function onLoad(executionContext) {
        return __awaiter(this, void 0, void 0, function () {
            var name_1;
            return __generator(this, function (_a) {
                try {
                    //#region form
                    // // ---------- FIELDS ----------
                    CommonD365.Fields.HideFields(executionContext, "crf6f_purchasedate");
                    //CommonD365.Fields.ShowFields(executionContext, "emailaddress1");
                    CommonD365.Fields.DisableFields(executionContext, "crf6f_make");
                    //CommonD365.Fields.EnableFields(executionContext, "telephone1");
                    CommonD365.Fields.SetRequired(executionContext, "crf6f_make");
                    CommonD365.Fields.SetOptional(executionContext, "crf6f_vehiclename");
                    CommonD365.Fields.SetRecommended(executionContext, "crf6f_year");
                    CommonD365.Fields.SetFocus(executionContext, "crf6f_make");
                    name_1 = CommonD365.Fields.GetValue(executionContext, "crf6f_vehiclename");
                    CommonD365.Fields.SetValue(executionContext, "description", name_1);
                    CommonD365.Fields.SetLabel(executionContext, "crf6f_iselectric", "Is Electric Vehicle?");
                    // ---------- SECTIONS ----------
                    CommonD365.Sections.HideSections(executionContext, "tab_2_section_3", "tab_2_section_4");
                    CommonD365.Sections.ShowSections(executionContext, "tab_2_section_5");
                    // ---------- TABS ----------
                    CommonD365.Tabs.HideTabs(executionContext, "tab_4");
                    CommonD365.Tabs.ShowTabs(executionContext, "tab_3");
                    CommonD365.Tabs.SetFocus(executionContext, "tab_2");
                    CommonD365.Tabs.SetLabel(executionContext, "tab_2", "Customer Details");
                    // ---------- Notifications ----------
                    CommonD365.Notifications.Form.ShowInfo(executionContext, "Record saved successfully.", "saveInfo");
                    CommonD365.Notifications.Form.ShowWarning(executionContext, "Missing important data.", "missingData");
                    CommonD365.Notifications.Form.ShowError(executionContext, "Unexpected error occurred.", "error1");
                    CommonD365.Notifications.Form.Clear(executionContext, "saveInfo");
                    //CommonD365.Notifications.Field.ShowError(executionContext, "emailaddress1", "Please enter a valid email.");
                    // CommonD365.Notifications.Field.ShowInfo(executionContext, "crf6f_vehiclename", "Enter your official first name.");
                    // CommonD365.Notifications.Field.Clear(executionContext, "emailaddress1");
                    // ---------- FORM ----------
                    // Save form
                    /*CommonD365.Form.save(executionContext);
                    
                    // Save and close form
                    await CommonD365.Form.saveAndClose(executionContext);
                    
                    // Refresh form data (without saving)
                    CommonD365.Form.refresh(executionContext);
                    
                    // Get form type
                    const formType = CommonD365.Form.getFormType(executionContext);
                    console.log("Form Type:", formType);
                    
                    // Refresh subgrids by name // mutliple subgrids
                    CommonD365.Form.refreshSubGrid(executionContext, "Contacts", "Opportunities");
                    
                    // Get all rows from a subgrid
                    const rows = CommonD365.Form.getSubGridRows(executionContext, "Contacts");
                    console.log("All subgrid rows:", rows);
                    
                    // Get selected rows from a subgrid
                    const selectedRows = CommonD365.Form.getSelectedSubGridRows(executionContext, "Contacts");
                    console.log("Selected subgrid rows:", selectedRows);
                    
                    // Prevent form from saving (useful in onSave event)
                    CommonD365.Form.preventSave(executionContext);
                    
                    // Refresh ribbon (e.g., to update button enable rules)
                    await CommonD365.Form.ribbonRefresh(executionContext);
                    
                    //---------- QUICK FORMS ----------
                    CommonD365.QuickForms.HideQuickForms(executionContext, "quickform_customer", "quickform_billing");
                    CommonD365.QuickForms.ShowQuickForms(executionContext, "quickform_customer");
                    //---------- ALERTS ----------
                    await CommonD365.Alerts.OpenAlertDialog(executionContext, "Record saved successfully.", "Success");
                    
                    const result = await CommonD365.Alerts.OpenConfirmDialog(executionContext, "Are you sure you want to delete this?", "Confirm Delete");
                    if (result.confirmed) {
                      console.log("User confirmed action.");
                    }
                    
                    //---------- User Info ----------
                    // Check if the user has a specific role
                    const isAgent = CommonD365.UserInfo.userHasRole(
                        executionContext,
                        CommonD365.Security.Roles.Agent
                    );
                    if (isAgent) {
                        console.log("current user is an Agent.");
                    } else {
                        console.log("Current user is not an Agent.");
                    }
                    
                    //  Check if the user has any one of multiple roles
                    const hasAnyRole = CommonD365.UserInfo.userHasAnyRole(
                        executionContext,
                        CommonD365.Security.Roles.Agent,
                        "jetManager",
                        "jetSupervisor"
                    );
                    if (hasAnyRole) {
                        console.log("User has at least one valid role.");
                    } else {
                        console.log("User does not have any of the specified roles.");
                    }
                    
                    // Get the current user's ID
                    const userId = CommonD365.UserInfo.getUserId(executionContext);
                    console.log("User ID:", userId);
                    
                    // Get the current user's full name
                    const userName = CommonD365.UserInfo.getUserName(executionContext);
                    console.log("User Name:", userName);
                    
                    // Get all user roles
                    const userRoles = CommonD365.UserInfo.getUserRoles(executionContext);
                    console.log("User Roles:", userRoles);
                    
                    
                    
                    //#endregion form
                    
                    //#region crud operations
                    // CREATE
                    const createData = {
                        name: "New Account Created from CommonD365.Data",
                        description: "Demo account record"
                    };
                    const createResponse = await CommonD365.Data.Create(executionContext,"account", createData);
                    console.log("Created Record ID:", createResponse?.id);
                    
                    // RETRIEVE
                    const retrieveResponse = await CommonD365.Data.Retrieve(executionContext,"account", "Guidneed to pass", "?$select=name,accountnumber");
                    console.log("Retrieved Record:", retrieveResponse);
                    
                    // UPDATE
                    const updateData = { name: "Updated Account Name" };
                    await CommonD365.Data.Update(executionContext,"account", "Guidneed to pass", updateData);
                    console.log("✏️ Record updated successfully");
                    
                    // RETRIEVE MULTIPLE
                    const multipleResponse = await CommonD365.Data.RetrieveMultiple(executionContext,"account", "?$select=name&$top=3");
                    console.log("📚 Top 3 Accounts:", multipleResponse?.entities);
                    
                    // EXECUTE GLOBAL ACTION
                    const globalActionResponse = await CommonD365.Data.ExecuteGlobalActionRequest(
                        executionContext,
                        "new_GlobalActionName",
                        "Sample Input Data"
                    );
                    console.log("Global Action Response:", globalActionResponse);
                    
                    // EXECUTE ENTITY-BOUND ACTION
                    const entityActionResponse = await CommonD365.Data.ExecuteEntityActionRequest(
                        executionContext,
                        "new_EntityBoundAction",
                        "account",
                        "Guidneed to pass",
                        "Input for Entity Action"
                    );
                    console.log("Entity Action Response:", entityActionResponse);
                    
                    // DELETE
                    await CommonD365.Data.Delete(executionContext,"account", "Guidneed to pass");
                    console.log("Record deleted successfully");
                    //#endregion
                        
                    */
                }
                catch (error) {
                    CommonD365.Generic.HandleError(executionContext, "Employee form onLoad error:" + error, "employeeform_onload");
                }
                return [2 /*return*/];
            });
        });
    }
    EmployeeForm.onLoad = onLoad;
})(EmployeeForm || (EmployeeForm = {}));
window.EmployeeForm = EmployeeForm;
