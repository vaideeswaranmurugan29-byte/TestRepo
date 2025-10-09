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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
var CommonD365;
(function (CommonD365) {
    //#region Generic
    var Generic;
    (function (Generic) {
        function IsNull(value, message, throwError) {
            if (message === void 0) { message = "Value is null or empty."; }
            if (throwError === void 0) { throwError = false; }
            var isNullOrUndefined = value === null || value === undefined;
            var isEmptyString = typeof value === "string" && value.trim().length === 0;
            if (isNullOrUndefined || isEmptyString) {
                if (throwError) {
                    throw new Error(message);
                }
                return true;
            }
            return false;
        }
        Generic.IsNull = IsNull;
        function HandleError(executionContext, error, uniqueId) {
            var _a, _b;
            try {
                var formContext = (_b = (_a = executionContext === null || executionContext === void 0 ? void 0 : executionContext.getFormContext) === null || _a === void 0 ? void 0 : _a.call(executionContext)) !== null && _b !== void 0 ? _b : executionContext;
                var message = error instanceof Error ? error.message : String(error);
                if (formContext && Notifications && Notifications.Form) {
                    Notifications.Form.ShowError(formContext, "An internal script error has occurred. " + message, uniqueId);
                }
                else {
                    console.error("Error [" + uniqueId + "]: " + message);
                }
            }
            catch (e) {
                console.error("Critical error in Generic.HandleError: " + e.message);
            }
        }
        Generic.HandleError = HandleError;
    })(Generic = CommonD365.Generic || (CommonD365.Generic = {}));
    //#endregion
    //#region Internal
    var Internal;
    (function (Internal) {
        function GetFormContext(executionContext) {
            try {
                if (executionContext && typeof executionContext.getFormContext === "function") {
                    return executionContext.getFormContext();
                }
                throw new Error("ExecutionContext is missing or invalid.");
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Internal.GetFormContext");
                throw error;
            }
        }
        Internal.GetFormContext = GetFormContext;
        function GetWebAPIContext(executionContext) {
            var _a, _b;
            try {
                if (!Xrm || !Xrm.Utility || !Xrm.Utility.getGlobalContext) {
                    throw new Error("Xrm context not available.");
                }
                var globalContext = Xrm.Utility.getGlobalContext();
                var clientState = ((_b = (_a = globalContext.client).getClientState) === null || _b === void 0 ? void 0 : _b.call(_a)) || "Online";
                if (clientState === "Online") {
                    return Xrm.WebApi.online;
                }
                if (Xrm.WebApi.offline) {
                    return Xrm.WebApi.offline;
                }
                throw new Error("Offline Web API not available.");
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Internal.GetWebAPIContext");
                throw error;
            }
        }
        Internal.GetWebAPIContext = GetWebAPIContext;
        function ShowHideControls(formContext, visible) {
            var controlNames = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                controlNames[_i - 2] = arguments[_i];
            }
            try {
                controlNames.forEach(function (name) {
                    if (!Generic.IsNull(name)) {
                        var control = formContext.getControl(name);
                        if (control)
                            control.setVisible(visible);
                    }
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.ShowHideControls");
            }
        }
        Internal.ShowHideControls = ShowHideControls;
        function EnableDisableControls(formContext, disabled) {
            var controlNames = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                controlNames[_i - 2] = arguments[_i];
            }
            try {
                controlNames.forEach(function (name) {
                    if (!Generic.IsNull(name)) {
                        var control = formContext.getControl(name);
                        if (control && typeof control.setDisabled === "function")
                            control.setDisabled(disabled);
                    }
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.EnableDisableControls");
            }
        }
        Internal.EnableDisableControls = EnableDisableControls;
        function SetFieldRequired(formContext, level) {
            var attributeNames = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                attributeNames[_i - 2] = arguments[_i];
            }
            try {
                attributeNames.forEach(function (name) {
                    if (!Generic.IsNull(name)) {
                        var attr = formContext.getAttribute(name);
                        if (attr)
                            attr.setRequiredLevel(level);
                    }
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.SetFieldRequired");
            }
        }
        Internal.SetFieldRequired = SetFieldRequired;
        function ShowHideSections(formContext, isVisible) {
            var sectionNames = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                sectionNames[_i - 2] = arguments[_i];
            }
            try {
                sectionNames.forEach(function (sectionName) {
                    formContext.ui.tabs.forEach(function (tab) {
                        var section = tab.sections.get(sectionName);
                        if (section)
                            section.setVisible(isVisible);
                    });
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.ShowHideSections");
            }
        }
        Internal.ShowHideSections = ShowHideSections;
        function ShowHideTabs(formContext, isVisible) {
            var tabNames = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                tabNames[_i - 2] = arguments[_i];
            }
            try {
                tabNames.forEach(function (tabName) {
                    if (!Generic.IsNull(tabName)) {
                        var tab = formContext.ui.tabs.get(tabName);
                        if (tab)
                            tab.setVisible(isVisible);
                    }
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.ShowHideTabs");
            }
        }
        Internal.ShowHideTabs = ShowHideTabs;
        function SetTabFocus(formContext, tabName) {
            try {
                if (Generic.IsNull(tabName))
                    return;
                var tab = formContext.ui.tabs.get(tabName);
                if (tab && typeof tab.setFocus === "function")
                    tab.setFocus();
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.SetTabFocus");
            }
        }
        Internal.SetTabFocus = SetTabFocus;
        function SetTabLabel(formContext, tabName, label) {
            try {
                if (Generic.IsNull(tabName) || Generic.IsNull(label))
                    return;
                var tab = formContext.ui.tabs.get(tabName);
                if (tab && typeof tab.setLabel === "function") {
                    tab.setLabel(label);
                }
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.SetTabLabel");
            }
        }
        Internal.SetTabLabel = SetTabLabel;
        function SetFocus(formContext, controlName) {
            try {
                var control = formContext.getControl(controlName);
                if (control && typeof control.setFocus === "function")
                    control.setFocus();
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.SetFocus");
            }
        }
        Internal.SetFocus = SetFocus;
        function GetValue(formContext, attributeName) {
            try {
                var attr = formContext.getAttribute(attributeName);
                return attr ? attr.getValue() : null;
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.GetValue");
                return null;
            }
        }
        Internal.GetValue = GetValue;
        function SetValue(formContext, attributeName, value) {
            try {
                var attr = formContext.getAttribute(attributeName);
                if (attr)
                    attr.setValue(value);
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.SetValue");
            }
        }
        Internal.SetValue = SetValue;
        function SetLabel(formContext, controlName, label) {
            try {
                var control = formContext.getControl(controlName);
                if (control && typeof control.setLabel === "function") {
                    control.setLabel(label);
                }
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.SetLabel");
            }
        }
        Internal.SetLabel = SetLabel;
        function RefreshSubgrids(formContext) {
            var gridNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                gridNames[_i - 1] = arguments[_i];
            }
            try {
                gridNames.forEach(function (gridName) {
                    var grid = formContext.getControl(gridName);
                    if (grid && typeof grid.refresh === "function") {
                        grid.refresh();
                    }
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.RefreshSubgrids");
            }
        }
        Internal.RefreshSubgrids = RefreshSubgrids;
        function GetGridContext(formContext, gridName) {
            try {
                var grid = formContext.getControl(gridName);
                return grid || null;
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.GetGridContext");
                return null;
            }
        }
        Internal.GetGridContext = GetGridContext;
        function ShowHideQuickForms(formContext, isVisible) {
            var quickFormNames = [];
            for (var _i = 2; _i < arguments.length; _i++) {
                quickFormNames[_i - 2] = arguments[_i];
            }
            try {
                if (Generic.IsNull(formContext, "Form context is null."))
                    return;
                if (quickFormNames.length === 0)
                    throw new Error("No quick form names provided.");
                quickFormNames.forEach(function (name) {
                    var control = formContext.getControl(name);
                    if (!control) {
                        console.warn("Quick form control '".concat(name, "' not found on the form."));
                        return;
                    }
                    // Dynamically show/hide any control (Quick Form included)
                    control.setVisible(isVisible);
                });
            }
            catch (error) {
                Generic.HandleError(formContext, error, "Internal.ShowHideQuickForms");
            }
        }
        Internal.ShowHideQuickForms = ShowHideQuickForms;
    })(Internal = CommonD365.Internal || (CommonD365.Internal = {}));
    //#endregion
    //#region Fields
    var Fields;
    (function (Fields) {
        function HideFields(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.ShowHideControls.apply(Internal, __spreadArray([formContext, false], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.HideFields");
            }
        }
        Fields.HideFields = HideFields;
        function ShowFields(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.ShowHideControls.apply(Internal, __spreadArray([formContext, true], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.ShowFields");
            }
        }
        Fields.ShowFields = ShowFields;
        function EnableFields(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.EnableDisableControls.apply(Internal, __spreadArray([formContext, false], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.EnableFields");
            }
        }
        Fields.EnableFields = EnableFields;
        function DisableFields(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.EnableDisableControls.apply(Internal, __spreadArray([formContext, true], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.DisableFields");
            }
        }
        Fields.DisableFields = DisableFields;
        function SetRequired(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetFieldRequired.apply(Internal, __spreadArray([formContext, "required"], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.SetRequired");
            }
        }
        Fields.SetRequired = SetRequired;
        function SetOptional(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetFieldRequired.apply(Internal, __spreadArray([formContext, "none"], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.SetOptional");
            }
        }
        Fields.SetOptional = SetOptional;
        function SetRecommended(executionContext) {
            var fieldNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                fieldNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetFieldRequired.apply(Internal, __spreadArray([formContext, "recommended"], fieldNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.SetRecommended");
            }
        }
        Fields.SetRecommended = SetRecommended;
        function SetFocus(executionContext, fieldName) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetFocus(formContext, fieldName);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.SetFocus");
            }
        }
        Fields.SetFocus = SetFocus;
        function GetValue(executionContext, fieldName) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                return Internal.GetValue(formContext, fieldName);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.GetValue");
                return null;
            }
        }
        Fields.GetValue = GetValue;
        function SetValue(executionContext, fieldName, value) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetValue(formContext, fieldName, value);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.SetValue");
            }
        }
        Fields.SetValue = SetValue;
        function SetLabel(executionContext, fieldName, label) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetLabel(formContext, fieldName, label);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Fields.SetLabel");
            }
        }
        Fields.SetLabel = SetLabel;
    })(Fields = CommonD365.Fields || (CommonD365.Fields = {}));
    //#endregion
    //#region Sections
    var Sections;
    (function (Sections) {
        function HideSections(executionContext) {
            var sectionNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                sectionNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.ShowHideSections.apply(Internal, __spreadArray([formContext, false], sectionNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Sections.HideSections");
            }
        }
        Sections.HideSections = HideSections;
        function ShowSections(executionContext) {
            var sectionNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                sectionNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.ShowHideSections.apply(Internal, __spreadArray([formContext, true], sectionNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Sections.ShowSections");
            }
        }
        Sections.ShowSections = ShowSections;
    })(Sections = CommonD365.Sections || (CommonD365.Sections = {}));
    //#endregion
    //#region Notifications
    var Notifications;
    (function (Notifications) {
        var Form;
        (function (Form) {
            function SetFormNotification(executionContext, message, uniqueId, notificationType) {
                try {
                    var formContext = Internal.GetFormContext(executionContext);
                    if (Generic.IsNull(formContext, "Form context is null."))
                        return;
                    if (Generic.IsNull(message, "Notification message is null."))
                        return;
                    if (Generic.IsNull(uniqueId, "Notification uniqueId is null."))
                        return;
                    if (Generic.IsNull(notificationType, "Notification type is null."))
                        return;
                    formContext.ui.setFormNotification(message, notificationType, uniqueId);
                }
                catch (e) {
                    console.error("Error in Notifications.Form.SetFormNotification: ".concat(e.message));
                }
            }
            Form.SetFormNotification = SetFormNotification;
            function ShowInfo(executionContext, message, uniqueId) {
                try {
                    SetFormNotification(executionContext, message, uniqueId, "INFO");
                }
                catch (e) {
                    console.error("Error in Notifications.Form.ShowInfo: ".concat(e.message));
                }
            }
            Form.ShowInfo = ShowInfo;
            function ShowWarning(executionContext, message, uniqueId) {
                try {
                    SetFormNotification(executionContext, message, uniqueId, "WARNING");
                }
                catch (e) {
                    console.error("Error in Notifications.Form.ShowWarning: ".concat(e.message));
                }
            }
            Form.ShowWarning = ShowWarning;
            function ShowError(executionContext, message, uniqueId) {
                try {
                    SetFormNotification(executionContext, message, uniqueId, "ERROR");
                }
                catch (e) {
                    console.error("Error in Notifications.Form.ShowError: ".concat(e.message));
                }
            }
            Form.ShowError = ShowError;
            function Clear(executionContext, uniqueId) {
                try {
                    var formContext = Internal.GetFormContext(executionContext);
                    if (Generic.IsNull(formContext, "Form context is null."))
                        return;
                    if (Generic.IsNull(uniqueId, "Notification uniqueId is null."))
                        return;
                    formContext.ui.clearFormNotification(uniqueId);
                }
                catch (e) {
                    console.error("Error in Notifications.Form.Clear: ".concat(e.message));
                }
            }
            Form.Clear = Clear;
        })(Form = Notifications.Form || (Notifications.Form = {}));
        var Field;
        (function (Field) {
            function GetControl(executionContext, fieldName) {
                var formContext = Internal.GetFormContext(executionContext);
                if (!formContext)
                    throw new Error("Form context not found.");
                var control = formContext.getControl(fieldName);
                if (!control)
                    throw new Error("Control '".concat(fieldName, "' not found."));
                // Verify it supports notifications
                var standardControl = control;
                if (!standardControl.setNotification || !standardControl.clearNotification) {
                    throw new Error("Control '".concat(fieldName, "' does not support notifications."));
                }
                return standardControl;
            }
            function SetNotification(executionContext, fieldName, message, uniqueId) {
                var control = GetControl(executionContext, fieldName);
                control.setNotification(message, uniqueId);
            }
            function ClearNotification(executionContext, fieldName, uniqueId) {
                var control = GetControl(executionContext, fieldName);
                if (uniqueId)
                    control.clearNotification(uniqueId);
                else
                    control.clearNotification();
            }
            function ShowError(executionContext, fieldName, message, uniqueId) {
                try {
                    var id = uniqueId || "".concat(fieldName, "_Error");
                    SetNotification(executionContext, fieldName, message, id);
                }
                catch (error) {
                    Generic.HandleError(executionContext, error, "Field.ShowError");
                }
            }
            Field.ShowError = ShowError;
            function ShowInfo(executionContext, fieldName, message, uniqueId) {
                try {
                    var id = uniqueId || "".concat(fieldName, "_Info");
                    SetNotification(executionContext, fieldName, message, id);
                }
                catch (error) {
                    Generic.HandleError(executionContext, error, "Field.ShowInfo");
                }
            }
            Field.ShowInfo = ShowInfo;
            function Clear(executionContext, fieldName, uniqueId) {
                try {
                    ClearNotification(executionContext, fieldName, uniqueId);
                }
                catch (error) {
                    Generic.HandleError(executionContext, error, "Field.Clear");
                }
            }
            Field.Clear = Clear;
        })(Field = Notifications.Field || (Notifications.Field = {}));
    })(Notifications = CommonD365.Notifications || (CommonD365.Notifications = {}));
    //#endregion
    //#region Tabs
    var Tabs;
    (function (Tabs) {
        function HideTabs(executionContext) {
            var tabNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                tabNames[_i - 1] = arguments[_i];
            }
            if (!executionContext || tabNames.length === 0) {
                Notifications.Form.ShowError(executionContext, "Execution context or tab names missing.", "Tabs");
                return;
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                if (Generic.IsNull(formContext, "Form context is null."))
                    return;
                Internal.ShowHideTabs.apply(Internal, __spreadArray([formContext, false], tabNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Tabs.HideTabs");
            }
        }
        Tabs.HideTabs = HideTabs;
        function ShowTabs(executionContext) {
            var tabNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                tabNames[_i - 1] = arguments[_i];
            }
            if (!executionContext || tabNames.length === 0) {
                Notifications.Form.ShowError(executionContext, "Execution context or tab names missing.", "Tabs");
                return;
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                if (Generic.IsNull(formContext, "Form context is null."))
                    return;
                Internal.ShowHideTabs.apply(Internal, __spreadArray([formContext, true], tabNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Tabs.ShowTabs");
            }
        }
        Tabs.ShowTabs = ShowTabs;
        function SetFocus(executionContext, tabName) {
            if (!executionContext || Generic.IsNull(tabName)) {
                Notifications.Form.ShowError(executionContext, "Execution context or tab name missing.", "Tabs");
                return;
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetTabFocus(formContext, tabName);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Tabs.SetFocus");
            }
        }
        Tabs.SetFocus = SetFocus;
        function SetLabel(executionContext, tabName, label) {
            if (!executionContext || Generic.IsNull(tabName) || Generic.IsNull(label)) {
                Notifications.Form.ShowError(executionContext, "Execution context, tab name, or label missing.", "Tabs");
                return;
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.SetTabLabel(formContext, tabName, label);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Tabs.SetLabel");
            }
        }
        Tabs.SetLabel = SetLabel;
    })(Tabs = CommonD365.Tabs || (CommonD365.Tabs = {}));
    //#endregion 
    //#region Form
    var Form;
    (function (Form) {
        function save(executionContext) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                if (!formContext)
                    throw new Error("Form context not found.");
                // Use explicit SaveOptions to satisfy typings
                var saveOptions = { saveMode: 1 }; // 1 = Save
                formContext.data.save(saveOptions);
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.save");
            }
        }
        Form.save = save;
        function saveAndClose(executionContext) {
            return __awaiter(this, void 0, void 0, function () {
                var formContext, saveOptions, error_1;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            formContext = Internal.GetFormContext(executionContext);
                            if (!formContext)
                                throw new Error("Form context not found.");
                            saveOptions = { saveMode: 1 };
                            return [4 /*yield*/, formContext.data.save(saveOptions)];
                        case 1:
                            _a.sent();
                            formContext.ui.close();
                            return [3 /*break*/, 3];
                        case 2:
                            error_1 = _a.sent();
                            Generic.HandleError(executionContext, error_1, "Form.saveAndClose");
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Form.saveAndClose = saveAndClose;
        function refresh(executionContext) {
            return __awaiter(this, void 0, void 0, function () {
                var formContext, error_2;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            formContext = Internal.GetFormContext(executionContext);
                            if (!formContext)
                                throw new Error("Form context not found.");
                            return [4 /*yield*/, formContext.data.refresh(false)];
                        case 1:
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_2 = _a.sent();
                            Generic.HandleError(executionContext, error_2, "Form.refresh");
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Form.refresh = refresh;
        function getFormType(executionContext) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                if (!formContext)
                    throw new Error("Form context not found.");
                return formContext.ui.getFormType();
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.getFormType");
                return null;
            }
        }
        Form.getFormType = getFormType;
        function refreshSubGrid(executionContext) {
            var gridNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                gridNames[_i - 1] = arguments[_i];
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                if (!formContext)
                    throw new Error("Form context not found.");
                Internal.RefreshSubgrids.apply(Internal, __spreadArray([formContext], gridNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.refreshSubGrid");
            }
        }
        Form.refreshSubGrid = refreshSubGrid;
        function getSubGridRows(executionContext, gridName) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                var gridContext = Internal.GetGridContext(formContext, gridName);
                if (!gridContext)
                    throw new Error("Grid '".concat(gridName, "' not found."));
                return gridContext.getGrid().getRows();
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.getSubGridRows");
                return null;
            }
        }
        Form.getSubGridRows = getSubGridRows;
        function getSelectedSubGridRows(executionContext, gridName) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                var gridContext = Internal.GetGridContext(formContext, gridName);
                if (!gridContext)
                    throw new Error("Grid '".concat(gridName, "' not found."));
                return gridContext.getGrid().getSelectedRows();
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.getSelectedSubGridRows");
                return null;
            }
        }
        Form.getSelectedSubGridRows = getSelectedSubGridRows;
        function preventSave(executionContext) {
            try {
                var ctx = executionContext;
                var eventArgs = typeof ctx.getEventArgs === "function" ? ctx.getEventArgs() : undefined;
                if (eventArgs && typeof eventArgs.preventDefault === "function") {
                    eventArgs.preventDefault();
                }
                else {
                    throw new Error("Event arguments not available — ensure this is called from onSave event.");
                }
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.preventSave");
            }
        }
        Form.preventSave = preventSave;
        function ribbonRefresh(executionContext) {
            try {
                var formContext = Internal.GetFormContext(executionContext);
                if (!formContext)
                    throw new Error("Form context not found.");
                formContext.ui.refreshRibbon();
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "Form.ribbonRefresh");
            }
        }
        Form.ribbonRefresh = ribbonRefresh;
    })(Form = CommonD365.Form || (CommonD365.Form = {}));
    //#endregion
    //#region quick form
    var QuickForms;
    (function (QuickForms) {
        function HideQuickForms(executionContext) {
            var quickFormNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                quickFormNames[_i - 1] = arguments[_i];
            }
            if (!executionContext || quickFormNames.length === 0) {
                Generic.HandleError(executionContext, new Error("Execution context or quickForm names missing."), "QuickForms.HideQuickForms");
                return;
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.ShowHideQuickForms.apply(Internal, __spreadArray([formContext, false], quickFormNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "QuickForms.HideQuickForms");
            }
        }
        QuickForms.HideQuickForms = HideQuickForms;
        function ShowQuickForms(executionContext) {
            var quickFormNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                quickFormNames[_i - 1] = arguments[_i];
            }
            if (!executionContext || quickFormNames.length === 0) {
                Generic.HandleError(executionContext, new Error("Execution context or quickForm names missing."), "QuickForms.ShowQuickForms");
                return;
            }
            try {
                var formContext = Internal.GetFormContext(executionContext);
                Internal.ShowHideQuickForms.apply(Internal, __spreadArray([formContext, true], quickFormNames, false));
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "QuickForms.ShowQuickForms");
            }
        }
        QuickForms.ShowQuickForms = ShowQuickForms;
    })(QuickForms = CommonD365.QuickForms || (CommonD365.QuickForms = {}));
    //#endregion
    //#region Alerts
    var Alerts;
    (function (Alerts) {
        function OpenAlertDialog(executionContext_1, message_1) {
            return __awaiter(this, arguments, void 0, function (executionContext, message, title, confirmLabel, height, width) {
                var alertStrings, alertOptions, error_3;
                if (title === void 0) { title = "Alert"; }
                if (confirmLabel === void 0) { confirmLabel = "OK"; }
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (Generic.IsNull(message, "Alert message is null."))
                                return [2 /*return*/];
                            alertStrings = {
                                text: message,
                                title: title
                            };
                            alertOptions = {
                                confirmButtonLabel: confirmLabel
                            };
                            if (height !== undefined)
                                alertOptions.height = height;
                            if (width !== undefined)
                                alertOptions.width = width;
                            return [4 /*yield*/, Xrm.Navigation.openAlertDialog(alertStrings, alertOptions)];
                        case 1:
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_3 = _a.sent();
                            Generic.HandleError(executionContext, error_3, "Alerts.OpenAlertDialog");
                            return [3 /*break*/, 3];
                        case 3: 
                        // Always return explicitly (to satisfy TS return rules)
                        return [2 /*return*/];
                    }
                });
            });
        }
        Alerts.OpenAlertDialog = OpenAlertDialog;
        function OpenConfirmDialog(executionContext_1, message_1) {
            return __awaiter(this, arguments, void 0, function (executionContext, message, title, confirmLabel, cancelLabel, height, width) {
                var confirmStrings, confirmOptions, result, error_4;
                var _a;
                if (title === void 0) { title = "Confirm Action"; }
                if (confirmLabel === void 0) { confirmLabel = "Yes"; }
                if (cancelLabel === void 0) { cancelLabel = "No"; }
                return __generator(this, function (_b) {
                    switch (_b.label) {
                        case 0:
                            _b.trys.push([0, 2, , 3]);
                            if (Generic.IsNull(message, "Confirm message is null.")) {
                                return [2 /*return*/, { confirmed: false }];
                            }
                            confirmStrings = {
                                text: message,
                                title: title
                            };
                            confirmOptions = {
                                confirmButtonLabel: confirmLabel,
                                cancelButtonLabel: cancelLabel
                            };
                            if (height !== undefined)
                                confirmOptions.height = height;
                            if (width !== undefined)
                                confirmOptions.width = width;
                            return [4 /*yield*/, Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions)];
                        case 1:
                            result = _b.sent();
                            return [2 /*return*/, (_a = result) !== null && _a !== void 0 ? _a : { confirmed: false }];
                        case 2:
                            error_4 = _b.sent();
                            Generic.HandleError(executionContext, error_4, "Alerts.OpenConfirmDialog");
                            return [2 /*return*/, { confirmed: false }];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Alerts.OpenConfirmDialog = OpenConfirmDialog;
    })(Alerts = CommonD365.Alerts || (CommonD365.Alerts = {}));
    //#endregion
    //#region Data (crud operations)
    var Data;
    (function (Data) {
        // CREATE
        function Create(executionContext, entityLogicalName, data) {
            return __awaiter(this, void 0, void 0, function () {
                var api, response, error_5;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!entityLogicalName)
                                throw new Error("entityLogicalName is required.");
                            if (!data)
                                throw new Error("data is required.");
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.createRecord(entityLogicalName, data)];
                        case 1:
                            response = _a.sent();
                            return [2 /*return*/, response];
                        case 2:
                            error_5 = _a.sent();
                            Generic.HandleError(executionContext, error_5, "Data.Create");
                            return [2 /*return*/, null];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.Create = Create;
        // RETRIEVE
        function Retrieve(executionContext, entityLogicalName, id, query) {
            return __awaiter(this, void 0, void 0, function () {
                var api, response, error_6;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!entityLogicalName)
                                throw new Error("entityLogicalName is required.");
                            if (!id)
                                throw new Error("id is required.");
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.retrieveRecord(entityLogicalName, id, query)];
                        case 1:
                            response = _a.sent();
                            return [2 /*return*/, response];
                        case 2:
                            error_6 = _a.sent();
                            Generic.HandleError(executionContext, error_6, "Data.Retrieve");
                            return [2 /*return*/, null];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.Retrieve = Retrieve;
        // UPDATE
        function Update(executionContext, entityLogicalName, id, data) {
            return __awaiter(this, void 0, void 0, function () {
                var api, error_7;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!entityLogicalName)
                                throw new Error("entityLogicalName is required.");
                            if (!id)
                                throw new Error("id is required.");
                            if (!data)
                                throw new Error("data is required.");
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.updateRecord(entityLogicalName, id, data)];
                        case 1:
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_7 = _a.sent();
                            Generic.HandleError(executionContext, error_7, "Data.Update");
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.Update = Update;
        // DELETE
        function Delete(executionContext, entityLogicalName, id) {
            return __awaiter(this, void 0, void 0, function () {
                var api, error_8;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!entityLogicalName)
                                throw new Error("entityLogicalName is required.");
                            if (!id)
                                throw new Error("id is required.");
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.deleteRecord(entityLogicalName, id)];
                        case 1:
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_8 = _a.sent();
                            Generic.HandleError(executionContext, error_8, "Data.Delete");
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.Delete = Delete;
        // RETRIEVE MULTIPLE
        function RetrieveMultiple(executionContext, entityLogicalName, query) {
            return __awaiter(this, void 0, void 0, function () {
                var api, response, error_9;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!entityLogicalName)
                                throw new Error("entityLogicalName is required.");
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.retrieveMultipleRecords(entityLogicalName, query || "")];
                        case 1:
                            response = _a.sent();
                            return [2 /*return*/, response];
                        case 2:
                            error_9 = _a.sent();
                            Generic.HandleError(executionContext, error_9, "Data.RetrieveMultiple");
                            return [2 /*return*/, null];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.RetrieveMultiple = RetrieveMultiple;
        // EXECUTE SINGLE REQUEST
        function Execute(executionContext, request) {
            return __awaiter(this, void 0, void 0, function () {
                var api, response, error_10;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!request)
                                throw new Error("request is required.");
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.execute(request)];
                        case 1:
                            response = _a.sent();
                            return [2 /*return*/, response];
                        case 2:
                            error_10 = _a.sent();
                            Generic.HandleError(executionContext, error_10, "Data.Execute");
                            return [2 /*return*/, null];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.Execute = Execute;
        // EXECUTE MULTIPLE REQUESTS
        function ExecuteMultiple(executionContext, requests) {
            return __awaiter(this, void 0, void 0, function () {
                var api_1, executions, error_11;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!Array.isArray(requests) || requests.length === 0) {
                                throw new Error("requests must be a non-empty array.");
                            }
                            api_1 = Internal.GetWebAPIContext();
                            executions = requests.map(function (req) {
                                return api_1.execute(req);
                            });
                            return [4 /*yield*/, Promise.all(executions)];
                        case 1: return [2 /*return*/, _a.sent()];
                        case 2:
                            error_11 = _a.sent();
                            Generic.HandleError(executionContext, error_11, "Data.ExecuteMultiple");
                            return [2 /*return*/, []];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.ExecuteMultiple = ExecuteMultiple;
        // EXECUTE GLOBAL ACTION REQUEST
        function ExecuteGlobalActionRequest(executionContext, actionName, data) {
            return __awaiter(this, void 0, void 0, function () {
                var request, api, response, error_12;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!actionName)
                                throw new Error("actionName is required.");
                            request = {
                                Input: data,
                                getMetadata: function () { return ({
                                    boundParameter: null,
                                    parameterTypes: {
                                        Input: { typeName: "Edm.String", structuralProperty: 1 }
                                    },
                                    operationType: 0,
                                    operationName: actionName
                                }); }
                            };
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.execute(request)];
                        case 1:
                            response = _a.sent();
                            return [2 /*return*/, response];
                        case 2:
                            error_12 = _a.sent();
                            Generic.HandleError(executionContext, error_12, "Data.ExecuteGlobalActionRequest");
                            return [2 /*return*/, null];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.ExecuteGlobalActionRequest = ExecuteGlobalActionRequest;
        // EXECUTE ENTITY ACTION REQUEST
        function ExecuteEntityActionRequest(executionContext, actionName, entityName, id, data) {
            return __awaiter(this, void 0, void 0, function () {
                var request, api, response, error_13;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            if (!actionName)
                                throw new Error("actionName is required.");
                            if (!entityName)
                                throw new Error("entityName is required.");
                            if (!id)
                                throw new Error("id is required.");
                            request = {
                                entity: { id: id, entityType: entityName },
                                Input: data,
                                getMetadata: function () { return ({
                                    boundParameter: "entity",
                                    parameterTypes: {
                                        entity: { typeName: "mscrm.".concat(entityName), structuralProperty: 5 },
                                        Input: { typeName: "Edm.String", structuralProperty: 1 }
                                    },
                                    operationType: 0,
                                    operationName: actionName
                                }); }
                            };
                            api = Internal.GetWebAPIContext();
                            return [4 /*yield*/, api.execute(request)];
                        case 1:
                            response = _a.sent();
                            return [2 /*return*/, response];
                        case 2:
                            error_13 = _a.sent();
                            Generic.HandleError(executionContext, error_13, "Data.ExecuteEntityActionRequest");
                            return [2 /*return*/, null];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        Data.ExecuteEntityActionRequest = ExecuteEntityActionRequest;
    })(Data = CommonD365.Data || (CommonD365.Data = {}));
    //#endregion
    //#region UserInfo
    var UserInfo;
    (function (UserInfo) {
        /** Checks if current user has a specific role */
        function userHasRole(executionContext, roleName) {
            try {
                if (!roleName)
                    throw new Error("Role name must be provided.");
                var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
                if (!userRoles || userRoles.getLength() === 0)
                    return false;
                for (var i = 0; i < userRoles.getLength(); i++) {
                    var role = userRoles.get(i);
                    var roleNameInSystem = role.name;
                    if (roleNameInSystem && roleNameInSystem.toLowerCase() === roleName.toLowerCase()) {
                        return true;
                    }
                }
                return false;
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "UserInfo.userHasRole");
                return false;
            }
        }
        UserInfo.userHasRole = userHasRole;
        /** Checks if user has any one of the given roles */
        function userHasAnyRole(executionContext) {
            var roleNames = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                roleNames[_i - 1] = arguments[_i];
            }
            try {
                if (roleNames.length === 0)
                    return false;
                var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
                if (!userRoles || userRoles.getLength() === 0)
                    return false;
                var lowerRoleNames = roleNames.map(function (r) { return r.toLowerCase(); });
                for (var i = 0; i < userRoles.getLength(); i++) {
                    var role = userRoles.get(i);
                    var roleNameInSystem = role.name;
                    if (roleNameInSystem && lowerRoleNames.indexOf(roleNameInSystem.toLowerCase()) !== -1) {
                        return true;
                    }
                }
                return false;
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "UserInfo.userHasAnyRole");
                return false;
            }
        }
        UserInfo.userHasAnyRole = userHasAnyRole;
        /** Gets current user ID (without braces) */
        function getUserId(executionContext) {
            try {
                var id = Xrm.Utility.getGlobalContext().userSettings.userId;
                return id ? id.replace(/[{}]/g, "") : null;
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "UserInfo.getUserId");
                return null;
            }
        }
        UserInfo.getUserId = getUserId;
        /** Gets current user full name */
        function getUserName(executionContext) {
            try {
                var name_1 = Xrm.Utility.getGlobalContext().userSettings.userName;
                return name_1 || null;
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "UserInfo.getUserName");
                return null;
            }
        }
        UserInfo.getUserName = getUserName;
        /** Returns list of all role names assigned to current user */
        function getUserRoles(executionContext) {
            try {
                var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
                var roleNames = [];
                if (!userRoles || userRoles.getLength() === 0)
                    return roleNames;
                for (var i = 0; i < userRoles.getLength(); i++) {
                    var role = userRoles.get(i);
                    var name_2 = role.name;
                    if (name_2)
                        roleNames.push(name_2);
                }
                return roleNames;
            }
            catch (error) {
                Generic.HandleError(executionContext, error, "UserInfo.getUserRoles");
                return [];
            }
        }
        UserInfo.getUserRoles = getUserRoles;
    })(UserInfo = CommonD365.UserInfo || (CommonD365.UserInfo = {}));
    //#endregion
    //#region Security Roles
    var Security;
    (function (Security) {
        var Roles;
        (function (Roles) {
            Roles.Agent = "jetAgent";
            Roles.Manager = "jetManager";
            Roles.Admin = "jetAdmin";
        })(Roles = Security.Roles || (Security.Roles = {}));
    })(Security = CommonD365.Security || (CommonD365.Security = {}));
    //#endregion 
})(CommonD365 || (CommonD365 = {}));
window.CommonD365 = CommonD365;
