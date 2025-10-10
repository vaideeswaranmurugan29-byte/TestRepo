
//#region Generic Utilities
export const Generic = {
  IsNull(value: any, message = "Value is null or empty.", throwError = false): boolean {
    const isNullOrUndefined = value === null || value === undefined;
    const isEmptyString = typeof value === "string" && value.trim().length === 0;

    if (isNullOrUndefined || isEmptyString) {
      if (throwError) {
        throw new Error(message);
      }
      return true;
    }
    return false;
  },
  HandleError(
    executionContext: Xrm.Events.EventContext,
    e: any,
    area: string,
    Notifications?: any
  ): void {
    const errorMessage = `An internal script error has occurred in ${area}. ${e?.message ?? e}`;
    console.error(errorMessage, e);

    try {
      // Notifications passed in from CommonD365 — avoids circular dependency
      const formNotifier = Notifications?.Form;

      if (formNotifier && typeof formNotifier.ShowError === "function") {
        formNotifier.ShowError(executionContext, errorMessage, area);
      } else {
        console.warn("Form notification handler not available; logged to console only.");
      }
    } catch (notifyError) {
      console.error("Failed to show CRM form notification:", notifyError);
    }
  },
};
//#endregion

//#region Internal Helpers
export const Internal = {
  GetFormContext(executionContext: Xrm.Events.EventContext): Xrm.FormContext {
    if (executionContext && typeof executionContext.getFormContext === "function") {
      return executionContext.getFormContext();
    }
    throw new Error("ExecutionContext is missing or invalid.");
  },
ShowHideSections(formContext: Xrm.FormContext, visible: boolean, ...sectionNames: string[]): void {
  formContext.ui.tabs.forEach((tab) => {
    tab.sections.forEach((section) => {
      if (sectionNames.includes(section.getName())) {
        section.setVisible(visible);
      }
    });
  });
},
ShowHideTabs(formContext: Xrm.FormContext, visible: boolean, ...tabNames: string[]): void {
  tabNames.forEach((name) => {
    if (!Generic.IsNull(name)) {
      const tab = formContext.ui.tabs.get(name);
      if (tab) {
        tab.setVisible(visible);
      }
    }
  });
},

SetTabFocus(formContext: Xrm.FormContext, tabName: string): void {
  const tab = formContext.ui.tabs.get(tabName);
  if (tab && typeof tab.setFocus === "function") {
    tab.setFocus();
  }
},
  ShowHideControls(formContext: Xrm.FormContext, visible: boolean, ...controlNames: string[]): void {
    controlNames.forEach((name) => {
      if (!Generic.IsNull(name)) {
        const control = formContext.getControl<Xrm.Controls.StandardControl>(name);
        if (control) control.setVisible(visible);
      }
    });
  },

  EnableDisableControls(formContext: Xrm.FormContext, disabled: boolean, ...controlNames: string[]): void {
    controlNames.forEach((name) => {
      if (!Generic.IsNull(name)) {
        const control = formContext.getControl<Xrm.Controls.StandardControl>(name);
        if (control && typeof control.setDisabled === "function") {
          control.setDisabled(disabled);
        }
      }
    });
  },

  SetFieldRequired(
    formContext: Xrm.FormContext,
    level: Xrm.Attributes.RequirementLevel,
    ...attributeNames: string[]
  ): void {
    attributeNames.forEach((name) => {
      if (!Generic.IsNull(name)) {
        const attr = formContext.getAttribute(name);
        if (attr) attr.setRequiredLevel(level);
      }
    });
  },

  SetFocus(formContext: Xrm.FormContext, controlName: string): void {
    const control = formContext.getControl<Xrm.Controls.StandardControl>(controlName);
    if (control && typeof control.setFocus === "function") {
      control.setFocus();
    }
  },

  GetValue<T = any>(formContext: Xrm.FormContext, attributeName: string): T | null {
    const attr = formContext.getAttribute(attributeName);
    return attr ? (attr.getValue() as T) : null;
  },

  SetValue<T = any>(formContext: Xrm.FormContext, attributeName: string, value: T): void {
    const attr = formContext.getAttribute(attributeName);
    if (attr && typeof (attr as any).setValue === "function") {
      (attr as any).setValue(value);
    }
  },

  SetLabel(formContext: Xrm.FormContext, controlName: string, label: string): void {
    const control = formContext.getControl(controlName);
    if (control && typeof control.setLabel === "function") {
      control.setLabel(label);
    }
  },

GetGridContext(formContext: Xrm.FormContext, gridName: string): Xrm.Controls.GridControl | null {
  const grid = formContext.getControl<Xrm.Controls.GridControl>(gridName);
  return grid || null;
},

RefreshSubgrids(formContext: Xrm.FormContext, ...gridNames: string[]): void {
  gridNames.forEach((gridName) => {
    const grid = formContext.getControl<Xrm.Controls.GridControl>(gridName);
    if (grid && grid.refresh) {
      grid.refresh();
    } else {
      console.warn(`Grid '${gridName}' not found or refresh not supported.`);
    }
  });
},
ShowHideQuickForms(formContext: Xrm.FormContext, visible: boolean, ...quickFormNames: string[]): void {
  quickFormNames.forEach((name) => {
    const control = formContext.getControl<Xrm.Controls.StandardControl>(name);
    if (control && typeof control.setVisible === 'function') {
      control.setVisible(visible);
    } else {
      console.warn(`Quick form '${name}' not found or does not support visibility.`);
    }
  });
},


};
//#endregion

//#region Field-level Helpers
export const Fields = {
  HideFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.ShowHideControls(formContext, false, ...fieldNames);
  },

  ShowFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.ShowHideControls(formContext, true, ...fieldNames);
  },

  EnableFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.EnableDisableControls(formContext, false, ...fieldNames);
  },

  DisableFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.EnableDisableControls(formContext, true, ...fieldNames);
  },

  SetRequired(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.SetFieldRequired(formContext, "required", ...fieldNames);
  },

  SetOptional(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.SetFieldRequired(formContext, "none", ...fieldNames);
  },

  SetRecommended(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.SetFieldRequired(formContext, "recommended", ...fieldNames);
  },

  SetFocus(executionContext: Xrm.Events.EventContext, fieldName: string): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.SetFocus(formContext, fieldName);
  },

  GetValue<T = any>(executionContext: Xrm.Events.EventContext, fieldName: string): T | null {
    const formContext = Internal.GetFormContext(executionContext);
    return Internal.GetValue<T>(formContext, fieldName);
  },

  SetValue<T = any>(executionContext: Xrm.Events.EventContext, fieldName: string, value: T): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.SetValue<T>(formContext, fieldName, value);
  },

  SetLabel(executionContext: Xrm.Events.EventContext, fieldName: string, label: string): void {
    const formContext = Internal.GetFormContext(executionContext);
    Internal.SetLabel(formContext, fieldName, label);
  },
};
//#endregion
//#region Section-level Helpers
export const Sections = {
  HideSections(executionContext: Xrm.Events.EventContext, ...sectionNames: string[]): void {
    if (!executionContext || sectionNames.length === 0) {
      console.error("Execution context or section names missing.");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideSections(formContext, false, ...sectionNames);
    } catch (e: any) {
      console.error("Error hiding sections:", e.message || e);
    }
  },

  ShowSections(executionContext: Xrm.Events.EventContext, ...sectionNames: string[]): void {
    if (!executionContext || sectionNames.length === 0) {
      console.error("Execution context or section names missing.");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideSections(formContext, true, ...sectionNames);
    } catch (e: any) {
      console.error("Error showing sections:", e.message || e);
    }
  },
};
//#endregion
//#region Tabs Helpers
export const Tabs = {
  HideTabs(executionContext: Xrm.Events.EventContext, ...tabNames: string[]): void {
    if (!executionContext || tabNames.length === 0) {
      console.error("Execution context or tab names missing.");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideTabs(formContext, false, ...tabNames);
    } catch (e: any) {
      if (Notifications?.Form) {
        Notifications.Form.ShowError(
          executionContext,
          "An internal script error has occurred. " + e.message,
          "Tabs"
        );
      } else {
        console.error("Error:", e.message);
      }
    }
  },

  ShowTabs(executionContext: Xrm.Events.EventContext, ...tabNames: string[]): void {
    if (!executionContext || tabNames.length === 0) {
      console.error("Execution context or tab names missing.");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideTabs(formContext, true, ...tabNames);
    } catch (e: any) {
      if (Notifications?.Form) {
        Notifications.Form.ShowError(
          executionContext,
          "An internal script error has occurred. " + e.message,
          "Tabs"
        );
      } else {
        console.error("Error:", e.message);
      }
    }
  },

  SetFocus(executionContext: Xrm.Events.EventContext, tabName: string): void {
    if (!executionContext || !tabName) {
      console.error("Execution context or tab name missing.");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetTabFocus(formContext, tabName);
    } catch (e: any) {
      if (Notifications?.Form) {
        Notifications.Form.ShowError(
          executionContext,
          "An internal script error has occurred. " + e.message,
          "Tabs"
        );
      } else {
        console.error("Error:", e.message);
      }
    }
  },

  SetLabel(executionContext: Xrm.Events.EventContext, tabName: string, value: string): void {
    if (!executionContext || !tabName || !value) {
      console.error("Execution context, tab name, or label missing.");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      const tab = formContext.ui.tabs.get(tabName);
      if (tab) {
        tab.setLabel(value);
      } else {
        console.warn(`Tab '${tabName}' not found on form.`);
      }
    } catch (e: any) {
      if (Notifications?.Form) {
        Notifications.Form.ShowError(
          executionContext,
          "An internal script error has occurred. " + e.message,
          "Tabs"
        );
      } else {
        console.error("Error:", e.message);
      }
    }
  },
};
//#endregion



//#region Notifications Helpers
export const Notifications = {
  Form: {
    SetFormNotification(
      executionContext: Xrm.Events.EventContext,
      message: string,
      uniqueId: string,
      notificationType: Xrm.FormNotificationLevel
    ): void {
      const formContext = Internal.GetFormContext(executionContext);
      if (Generic.IsNull(formContext, "Form context is null.")) return;
      if (Generic.IsNull(message, "Notification message is null.")) return;
      if (Generic.IsNull(uniqueId, "Notification uniqueId is null.")) return;
      if (Generic.IsNull(notificationType, "Notification type is null.")) return;

      try {
        formContext.ui.setFormNotification(message, notificationType, uniqueId);
      } catch (e: any) {
        console.error("Failed to set form notification:", e.message);
      }
    },

    ShowInfo(executionContext: Xrm.Events.EventContext, message: string, uniqueId: string): void {
      this.SetFormNotification(executionContext, message, uniqueId, "INFO");
    },

    ShowWarning(executionContext: Xrm.Events.EventContext, message: string, uniqueId: string): void {
      this.SetFormNotification(executionContext, message, uniqueId, "WARNING");
    },

    ShowError(executionContext: Xrm.Events.EventContext, message: string, uniqueId: string): void {
      this.SetFormNotification(executionContext, message, uniqueId, "ERROR");
    },

    Clear(executionContext: Xrm.Events.EventContext, uniqueId: string): void {
      const formContext = Internal.GetFormContext(executionContext);
      if (Generic.IsNull(formContext, "Form context is null.")) return;
      if (Generic.IsNull(uniqueId, "Notification uniqueId is null.")) return;

      try {
        formContext.ui.clearFormNotification(uniqueId);
      } catch (e: any) {
        console.error("Failed to clear form notification:", e.message);
      }
    },
  },
};
//#endregion
//#region Form-level Helpers
export const Form = {
  save(executionContext: Xrm.Events.EventContext): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      formContext.data.save();
    } catch (e) {
      Generic.HandleError(executionContext, e, "Form", Notifications);
    }
  },

  saveAndClose(executionContext: Xrm.Events.EventContext): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      formContext.data.save().then(() => formContext.ui.close());
    } catch (e) {
      Generic.HandleError(executionContext, e, "Form", Notifications);
    }
  },

 refresh: async function (executionContext: Xrm.Events.EventContext) {
  try {
    const formContext = CommonD365.Internal.GetFormContext(executionContext);
    if (!formContext) throw new Error("Form context not found.");

    // Refresh without saving
    await formContext.data.refresh(false);

    console.log("Form data refreshed successfully.");
  } catch (e: any) {
    Generic.HandleError(executionContext, e, "Form");
  }
},


  getFormType(executionContext: Xrm.Events.EventContext): number | null {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      return formContext.ui.getFormType();
    } catch (e) {
      Generic.HandleError(executionContext, e, "Form", Notifications);
      return null;
    }
  },

  refreshSubGrid(executionContext: Xrm.Events.EventContext, ...gridNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      Internal.RefreshSubgrids(formContext, ...gridNames);
    } catch (e) {
      Generic.HandleError(executionContext, e, "Form", Notifications);
    }
  },

  getSubGridRows(executionContext: Xrm.Events.EventContext, gridName: string) {
  try {
    const formContext = CommonD365.Internal.GetFormContext(executionContext);
    const gridContext = CommonD365.Internal.GetGridContext(formContext, gridName);

    if (!gridContext) throw new Error(`Grid '${gridName}' not found.`);
    return gridContext.getGrid().getRows();
  } catch (e: any) {
    CommonD365.Generic.HandleError(executionContext, e, "Form");
    return null;
  }
},

getSelectedSubGridRows(executionContext: Xrm.Events.EventContext, gridName: string) {
  try {
    const formContext = CommonD365.Internal.GetFormContext(executionContext);
    const gridContext = CommonD365.Internal.GetGridContext(formContext, gridName);

    if (!gridContext) throw new Error(`Grid '${gridName}' not found.`);
    return gridContext.getGrid().getSelectedRows(); // ✅ Works fine
  } catch (e: any) {
    CommonD365.Generic.HandleError(executionContext, e, "Form");
    return null;
  }
},


  preventSave(executionContext: Xrm.Events.EventContext): void {
  try {
    const eventArgs = (executionContext as Xrm.Events.SaveEventContext).getEventArgs?.();
    if (eventArgs) {
      eventArgs.preventDefault();
    } else {
      throw new Error("Event arguments not available — ensure this is called from onSave event.");
    }
  } catch (e) {
    Generic.HandleError(executionContext, e, "Form", Notifications);
  }
},


  ribbonRefresh(executionContext: Xrm.Events.EventContext): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      formContext.ui.refreshRibbon();
    } catch (e) {
      Generic.HandleError(executionContext, e, "Form", Notifications);
    }
  },
};
//#endregion
//region QuickForm Helpers
export const QuickForm = {
  HideQuickForms(executionContext: Xrm.Events.EventContext, ...quickFormNames: string[]): void {
    if (!executionContext || quickFormNames.length === 0) {
      console.error('Execution context or quickForm names missing.');
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideQuickForms(formContext, false, ...quickFormNames);
    } catch (e: any) {
      Notifications?.Form?.ShowError(
        executionContext,
        'An internal script error has occurred. ' + e.message,
        'QuickForms'
      );
    }
  },

  ShowQuickForms(executionContext: Xrm.Events.EventContext, ...quickFormNames: string[]): void {
    if (!executionContext || quickFormNames.length === 0) {
      console.error('Execution context or quickForm names missing.');
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideQuickForms(formContext, true, ...quickFormNames);
    } catch (e: any) {
      Notifications?.Form?.ShowError(
        executionContext,
        'An internal script error has occurred. ' + e.message,
        'QuickForms'
      );
    }
  },
};
// #endregion

//Alerts Helpers
export const Alerts = {
  OpenAlertDialog: (
    executionContext: Xrm.Events.EventContext,
    message: string,
    title: string = "Alert",
    confirmLabel: string = "OK",
    height: number = 200,
    width: number = 400
  ): void => {
    try {
      if (Generic.IsNull(message, "Alert message is null.")) return;

      const alertStrings = {
        text: message,
        title,
      };

      const alertOptions = {
        height,
        width,
        confirmButtonLabel: confirmLabel,
      };

      Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).catch((e) => {
        Generic.HandleError(executionContext, e, "Alerts");
      });
    } catch (e: any) {
      Generic.HandleError(executionContext, e, "Alerts");
    }
  },

  OpenConfirmDialog: (
    executionContext: Xrm.Events.EventContext,
    message: string,
    title: string = "Confirm",
    confirmLabel: string = "Yes",
    cancelLabel: string = "No",
    height: number = 200,
    width: number = 400
  ): void => {
    try {
      if (Generic.IsNull(message, "Confirm message is null.")) return;

      const confirmStrings = {
        text: message,
        title,
      };

      const confirmOptions = {
        height,
        width,
        confirmButtonLabel: confirmLabel,
        cancelButtonLabel: cancelLabel,
      };

      Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions)
        .then((result) => {
          if (result.confirmed) {
            console.log("User confirmed action.");
          } else {
            console.log("User canceled action.");
          }
        })
        .catch((e) => {
          Generic.HandleError(executionContext, e, "Alerts");
        });
    } catch (e: any) {
      Generic.HandleError(executionContext, e, "Alerts");
    }
  },
};
//endregion

//region User Info Helpers
export const UserInfo = {
  /**
   * Checks if the current user has a specific role.
   */
  userHasRole(roleName: string): boolean {
    try {
      const userSettings = Xrm.Utility.getGlobalContext().userSettings;
      const userRoles = userSettings.roles;

      if (!roleName || !userRoles || userRoles.getLength() === 0) {
        return false;
      }

      for (let i = 0; i < userRoles.getLength(); i++) {
        const role = userRoles.get(i);
        if (role?.name?.toLowerCase() === roleName.toLowerCase()) {
          return true;
        }
      }

      return false;
    } catch (error) {
      console.error("Error in userHasRole:", error);
      return false;
    }
  },

  /**
   * Checks if the current user has any of the specified roles.
   */
  userHasAnyRole(...roleNames: string[]): boolean {
    try {
      if (roleNames.length === 0) {
        return false;
      }

      const userSettings = Xrm.Utility.getGlobalContext().userSettings;
      const userRoles = userSettings.roles;

      if (!userRoles || userRoles.getLength() === 0) {
        return false;
      }

      const lowerRoleNames = roleNames.map(r => r.toLowerCase());

      for (let i = 0; i < userRoles.getLength(); i++) {
        const role = userRoles.get(i);
        if (role?.name && lowerRoleNames.includes(role.name.toLowerCase())) {
          return true;
        }
      }

      return false;
    } catch (error) {
      console.error("Error in userHasAnyRole:", error);
      return false;
    }
  },

  /**
   * Gets the current user's ID (without curly braces).
   */
  getUserId(): string | null {
    try {
      const userId = Xrm.Utility.getGlobalContext().userSettings.userId;
      return userId ? userId.replace(/[{}]/g, "") : null;
    } catch (error) {
      console.error("Error in getUserId:", error);
      return null;
    }
  },

  /**
   * Gets the current user's full name.
   */
  getUserName(): string | null {
    try {
      return Xrm.Utility.getGlobalContext().userSettings.userName || null;
    } catch (error) {
      console.error("Error in getUserName:", error);
      return null;
    }
  },

  /**
   * Gets the list of current user’s role names.
   */
  getUserRoles(): string[] {
    try {
      const roles = Xrm.Utility.getGlobalContext().userSettings.roles;
      const roleNames: string[] = [];

      for (let i = 0; i < roles.getLength(); i++) {
        const role = roles.get(i);
        if (role?.name) {
          roleNames.push(role.name);
        }
      }

      return roleNames;
    } catch (error) {
      console.error("Error in getUserRoles:", error);
      return [];
    }
  },
};
//endregion


//#region Data (Web API) Helpers
export const Data = {
  /**
   * Create a new record.
   */
  async Create(entityLogicalName: string, data: any) {
    if (!entityLogicalName) throw new Error("entityLogicalName is required.");
    if (!data) throw new Error("data is required.");

    const api = Xrm.WebApi.online;
    return api.createRecord(entityLogicalName, data);
  },

  /**
   * Retrieve a record by ID.
   */
  async Retrieve(entityLogicalName: string, id: string, query?: string) {
    if (!entityLogicalName) throw new Error("entityLogicalName is required.");
    if (!id) throw new Error("id is required.");

    const api = Xrm.WebApi.online;
    return api.retrieveRecord(entityLogicalName, id, query);
  },

  /**
   * Update an existing record.
   */
  async Update(entityLogicalName: string, id: string, data: any) {
    if (!entityLogicalName) throw new Error("entityLogicalName is required.");
    if (!id) throw new Error("id is required.");
    if (!data) throw new Error("data is required.");

    const api = Xrm.WebApi.online;
    return api.updateRecord(entityLogicalName, id, data);
  },

  /**
   * Delete a record.
   */
  async Delete(entityLogicalName: string, id: string) {
    if (!entityLogicalName) throw new Error("entityLogicalName is required.");
    if (!id) throw new Error("id is required.");

    const api = Xrm.WebApi.online;
    return api.deleteRecord(entityLogicalName, id);
  },

  /**
   * Retrieve multiple records using a query.
   */
  async RetrieveMultiple(entityLogicalName: string, query?: string) {
    if (!entityLogicalName) throw new Error("entityLogicalName is required.");

    const api = Xrm.WebApi.online;
    return api.retrieveMultipleRecords(entityLogicalName, query || "");
  },

  /**
   * Execute a single Web API request (custom or standard).
   */
  async Execute(request: any) {
    if (!request) throw new Error("request is required.");

    const api = Xrm.WebApi.online;
    return api.execute(request);
  },

  /**
   * Execute multiple Web API requests in parallel.
   */
  async ExecuteMultiple(requests: any[]) {
    if (!Array.isArray(requests) || requests.length === 0)
      throw new Error("requests must be a non-empty array.");

    const api = Xrm.WebApi.online;
    const executions = requests.map((req) => api.execute(req));
    return Promise.all(executions);
  },

  /**
   * Execute a global (unbound) custom action.
   */
  async ExecuteGlobalActionRequest(actionName: string, data?: any) {
    if (!actionName) throw new Error("actionName is required.");

    const request = {
      Input: data ?? "",
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {
          Input: {
            typeName: "Edm.String",
            structuralProperty: 1,
          },
        },
        operationType: 0,
        operationName: actionName,
      }),
    };

    const api = Xrm.WebApi.online;
    return api.execute(request);
  },

  /**
   * Execute an entity-bound custom action.
   */
  async ExecuteEntityActionRequest(
    actionName: string,
    entityName: string,
    id: string,
    data?: any
  ) {
    if (!actionName) throw new Error("actionName is required.");
    if (!entityName) throw new Error("entityName is required.");
    if (!id) throw new Error("id is required.");

    const request = {
      entity: { id, entityType: entityName },
      Input: data ?? "",
      getMetadata: () => ({
        boundParameter: "entity",
        parameterTypes: {
          entity: {
            typeName: `mscrm.${entityName}`,
            structuralProperty: 5,
          },
          Input: {
            typeName: "Edm.String",
            structuralProperty: 1,
          },
        },
        operationType: 0,
        operationName: actionName,
      }),
    };

    const api = Xrm.WebApi.online;
    return api.execute(request);
  },
};
//#endregion

//#region security helpers
export const Security = {
  Roles: {
    Agent: "jetAgent",
    Manager: "jetManager",
    Admin: "System Administrator",
  },
};
//#endregion
//#region Export Root
const CommonD365 = {
  Generic,
  Internal,
  Fields,
  Sections,
  Tabs,
  Notifications,
  Form,
 QuickForm,
 Alerts,
 UserInfo,
 Data,
 Security
};

export default CommonD365;
//#endregion
