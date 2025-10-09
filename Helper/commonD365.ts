namespace CommonD365 {

  //#region Generic
export namespace Generic {
  export function IsNull(
    value: any,
    message = "Value is null or empty.",
    throwError = false
  ): boolean {
    const isNullOrUndefined = value === null || value === undefined;
    const isEmptyString = typeof value === "string" && value.trim().length === 0;

    if (isNullOrUndefined || isEmptyString) {
      if (throwError) {
        throw new Error(message);
      }
      return true;
    }
    return false;
  }

  export function HandleError(
    executionContext: Xrm.Events.EventContext | Xrm.FormContext,
    error: any,
    uniqueId: string
  ): void {
    try {
      const formContext =
        (executionContext as any)?.getFormContext?.() ?? executionContext;
      const message =
        error instanceof Error ? error.message : String(error);

      if (formContext && Notifications && Notifications.Form) {
        Notifications.Form.ShowError(
          formContext as any,
          "An internal script error has occurred. " + message,
          uniqueId
        );
      } else {
        console.error("Error [" + uniqueId + "]: " + message);
      }
    } catch (e: any) {
      console.error("Critical error in Generic.HandleError: " + e.message);
    }
  }
}


//#endregion
//#region Internal
 export namespace Internal {

  export function GetFormContext(executionContext: Xrm.Events.EventContext): Xrm.FormContext {
    try {
      if (executionContext && typeof executionContext.getFormContext === "function") {
        return executionContext.getFormContext();
      }
      throw new Error("ExecutionContext is missing or invalid.");
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Internal.GetFormContext");
      throw error;
    }
  }
export function GetWebAPIContext(executionContext?: Xrm.Events.EventContext) {
    try {
        if (!Xrm || !Xrm.Utility || !Xrm.Utility.getGlobalContext) {
            throw new Error("Xrm context not available.");
        }

        const globalContext = Xrm.Utility.getGlobalContext();
        const clientState = globalContext.client.getClientState?.() || "Online";

        if (clientState === "Online") {
            return Xrm.WebApi.online;
        }

        if (Xrm.WebApi.offline) {
            return Xrm.WebApi.offline;
        }

        throw new Error("Offline Web API not available.");
    } catch (error: any) {
        Generic.HandleError(executionContext as Xrm.Events.EventContext, error, "Internal.GetWebAPIContext");
        throw error;
    }
}
  export function ShowHideControls(formContext: Xrm.FormContext, visible: boolean, ...controlNames: string[]): void {
    try {
      controlNames.forEach(name => {
        if (!Generic.IsNull(name)) {
          const control = formContext.getControl<Xrm.Controls.StandardControl>(name);
          if (control) control.setVisible(visible);
        }
      });
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.ShowHideControls");
    }
  }

  export function EnableDisableControls(formContext: Xrm.FormContext, disabled: boolean, ...controlNames: string[]): void {
    try {
      controlNames.forEach(name => {
        if (!Generic.IsNull(name)) {
          const control = formContext.getControl<Xrm.Controls.StandardControl>(name);
          if (control && typeof control.setDisabled === "function") control.setDisabled(disabled);
        }
      });
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.EnableDisableControls");
    }
  }

  export function SetFieldRequired(formContext: Xrm.FormContext, level: Xrm.Attributes.RequirementLevel, ...attributeNames: string[]): void {
    try {
      attributeNames.forEach(name => {
        if (!Generic.IsNull(name)) {
          const attr = formContext.getAttribute(name);
          if (attr) attr.setRequiredLevel(level);
        }
      });
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.SetFieldRequired");
    }
  }

  export function ShowHideSections(formContext: Xrm.FormContext, isVisible: boolean, ...sectionNames: string[]): void {
    try {
      sectionNames.forEach(sectionName => {
        formContext.ui.tabs.forEach(tab => {
          const section = tab.sections.get(sectionName);
          if (section) section.setVisible(isVisible);
        });
      });
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.ShowHideSections");
    }
  }

  export function ShowHideTabs(formContext: Xrm.FormContext, isVisible: boolean, ...tabNames: string[]): void {
    try {
      tabNames.forEach(tabName => {
        if (!Generic.IsNull(tabName)) {
          const tab = formContext.ui.tabs.get(tabName);
          if (tab) tab.setVisible(isVisible);
        }
      });
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.ShowHideTabs");
    }
  }

  export function SetTabFocus(formContext: Xrm.FormContext, tabName: string): void {
    try {
      if (Generic.IsNull(tabName)) return;
      const tab = formContext.ui.tabs.get(tabName);
      if (tab && typeof tab.setFocus === "function") tab.setFocus();
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.SetTabFocus");
    }
  }

  export function SetTabLabel(formContext: Xrm.FormContext, tabName: string, label: string): void {
    try {
      if (Generic.IsNull(tabName) || Generic.IsNull(label)) return;
      const tab = formContext.ui.tabs.get(tabName);
      if (tab && typeof (tab as any).setLabel === "function") {
        (tab as any).setLabel(label);
      }
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.SetTabLabel");
    }
  }

  export function SetFocus(formContext: Xrm.FormContext, controlName: string): void {
    try {
      const control = formContext.getControl<Xrm.Controls.StandardControl>(controlName);
      if (control && typeof control.setFocus === "function") control.setFocus();
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.SetFocus");
    }
  }

  export function GetValue<T extends Xrm.Attributes.Attribute<any> = Xrm.Attributes.Attribute<any>>(
    formContext: Xrm.FormContext,
    attributeName: string
  ): ReturnType<T["getValue"]> | null {
    try {
      const attr = formContext.getAttribute<T>(attributeName);
      return attr ? attr.getValue() : null;
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.GetValue");
      return null;
    }
  }

  export function SetValue<T extends Xrm.Attributes.Attribute<any> = Xrm.Attributes.Attribute<any>>(
    formContext: Xrm.FormContext,
    attributeName: string,
    value: ReturnType<T["getValue"]>
  ): void {
    try {
      const attr = formContext.getAttribute<T>(attributeName);
      if (attr) attr.setValue(value);
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.SetValue");
    }
  }

  export function SetLabel(formContext: Xrm.FormContext, controlName: string, label: string): void {
    try {
      const control = formContext.getControl(controlName);
      if (control && typeof control.setLabel === "function") {
        control.setLabel(label);
      }
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.SetLabel");
    }
  }
  export function RefreshSubgrids(formContext: Xrm.FormContext, ...gridNames: string[]): void {
  try {
    gridNames.forEach(gridName => {
      const grid = formContext.getControl<Xrm.Controls.GridControl>(gridName);
      if (grid && typeof grid.refresh === "function") {
        grid.refresh();
      }
    });
  } catch (error: any) {
    Generic.HandleError(formContext as any, error, "Internal.RefreshSubgrids");
  }
}

export function GetGridContext(formContext: Xrm.FormContext, gridName: string): Xrm.Controls.GridControl | null {
  try {
    const grid = formContext.getControl<Xrm.Controls.GridControl>(gridName);
    return grid || null;
  } catch (error: any) {
    Generic.HandleError(formContext as any, error, "Internal.GetGridContext");
    return null;
  }
}
 export function ShowHideQuickForms(formContext: Xrm.FormContext, isVisible: boolean, ...quickFormNames: string[]): void {
    try {
      if (Generic.IsNull(formContext, "Form context is null.")) return;
      if (quickFormNames.length === 0) throw new Error("No quick form names provided.");

      quickFormNames.forEach((name) => {
        const control = formContext.getControl(name);

        if (!control) {
          console.warn(`Quick form control '${name}' not found on the form.`);
          return;
        }

        // Dynamically show/hide any control (Quick Form included)
        (control as Xrm.Controls.StandardControl).setVisible(isVisible);
      });
    } catch (error: any) {
      Generic.HandleError(formContext, error, "Internal.ShowHideQuickForms");
    }
  }

}

//#endregion
  
//#region Fields
 export namespace Fields {

  export function HideFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideControls(formContext, false, ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.HideFields");
    }
  }

  export function ShowFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideControls(formContext, true, ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.ShowFields");
    }
  }

  export function EnableFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.EnableDisableControls(formContext, false, ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.EnableFields");
    }
  }

  export function DisableFields(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.EnableDisableControls(formContext, true, ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.DisableFields");
    }
  }

  export function SetRequired(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetFieldRequired(formContext, "required", ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.SetRequired");
    }
  }

  export function SetOptional(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetFieldRequired(formContext, "none", ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.SetOptional");
    }
  }

  export function SetRecommended(executionContext: Xrm.Events.EventContext, ...fieldNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetFieldRequired(formContext, "recommended", ...fieldNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.SetRecommended");
    }
  }

  export function SetFocus(executionContext: Xrm.Events.EventContext, fieldName: string): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetFocus(formContext, fieldName);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.SetFocus");
    }
  }

  export function GetValue<T = any>(executionContext: Xrm.Events.EventContext, fieldName: string): T | null {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      return Internal.GetValue(formContext, fieldName) as T | null;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.GetValue");
      return null;
    }
  }

  export function SetValue<T = any>(executionContext: Xrm.Events.EventContext, fieldName: string, value: T): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetValue(formContext, fieldName, value);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.SetValue");
    }
  }

  export function SetLabel(executionContext: Xrm.Events.EventContext, fieldName: string, label: string): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetLabel(formContext, fieldName, label);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Fields.SetLabel");
    }
  }

}

  //#endregion

//#region Sections
 export namespace Sections {

  export function HideSections(executionContext: Xrm.Events.EventContext, ...sectionNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideSections(formContext, false, ...sectionNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Sections.HideSections");
    }
  }

  export function ShowSections(executionContext: Xrm.Events.EventContext, ...sectionNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideSections(formContext, true, ...sectionNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Sections.ShowSections");
    }
  }

}

//#endregion
//#region Notifications
export namespace Notifications {

  export namespace Form {

    export function SetFormNotification(
      executionContext: Xrm.Events.EventContext,
      message: string,
      uniqueId: string,
      notificationType: Xrm.FormNotificationLevel
    ): void {
      try {
        const formContext = Internal.GetFormContext(executionContext);

        if (Generic.IsNull(formContext, "Form context is null.")) return;
        if (Generic.IsNull(message, "Notification message is null.")) return;
        if (Generic.IsNull(uniqueId, "Notification uniqueId is null.")) return;
        if (Generic.IsNull(notificationType, "Notification type is null.")) return;

        formContext.ui.setFormNotification(message, notificationType, uniqueId);
      } catch (e: any) {
        console.error(`Error in Notifications.Form.SetFormNotification: ${e.message}`);
      }
    }

    export function ShowInfo(
      executionContext: Xrm.Events.EventContext,
      message: string,
      uniqueId: string
    ): void {
      try {
        SetFormNotification(executionContext, message, uniqueId, "INFO");
      } catch (e: any) {
        console.error(`Error in Notifications.Form.ShowInfo: ${e.message}`);
      }
    }

    export function ShowWarning(
      executionContext: Xrm.Events.EventContext,
      message: string,
      uniqueId: string
    ): void {
      try {
        SetFormNotification(executionContext, message, uniqueId, "WARNING");
      } catch (e: any) {
        console.error(`Error in Notifications.Form.ShowWarning: ${e.message}`);
      }
    }

    export function ShowError(
      executionContext: Xrm.Events.EventContext,
      message: string,
      uniqueId: string
    ): void {
      try {
        SetFormNotification(executionContext, message, uniqueId, "ERROR");
      } catch (e: any) {
        console.error(`Error in Notifications.Form.ShowError: ${e.message}`);
      }
    }

    export function Clear(
      executionContext: Xrm.Events.EventContext,
      uniqueId: string
    ): void {
      try {
        const formContext = Internal.GetFormContext(executionContext);

        if (Generic.IsNull(formContext, "Form context is null.")) return;
        if (Generic.IsNull(uniqueId, "Notification uniqueId is null.")) return;

        formContext.ui.clearFormNotification(uniqueId);
      } catch (e: any) {
        console.error(`Error in Notifications.Form.Clear: ${e.message}`);
      }
    }
  }
export namespace Field {

  function GetControl(
    executionContext: Xrm.Events.EventContext,
    fieldName: string
  ): Xrm.Controls.StandardControl {
    const formContext = Internal.GetFormContext(executionContext);
    if (!formContext) throw new Error("Form context not found.");

    const control = formContext.getControl(fieldName);
    if (!control) throw new Error(`Control '${fieldName}' not found.`);

    // Verify it supports notifications
    const standardControl = control as Xrm.Controls.StandardControl;
    if (!standardControl.setNotification || !standardControl.clearNotification) {
      throw new Error(`Control '${fieldName}' does not support notifications.`);
    }

    return standardControl;
  }

  function SetNotification(
    executionContext: Xrm.Events.EventContext,
    fieldName: string,
    message: string,
    uniqueId: string
  ): void {
    const control = GetControl(executionContext, fieldName);
    control.setNotification(message, uniqueId);
  }

  function ClearNotification(
    executionContext: Xrm.Events.EventContext,
    fieldName: string,
    uniqueId?: string
  ): void {
    const control = GetControl(executionContext, fieldName);
    if (uniqueId) control.clearNotification(uniqueId);
    else control.clearNotification();
  }

  export function ShowError(
    executionContext: Xrm.Events.EventContext,
    fieldName: string,
    message: string,
    uniqueId?: string
  ): void {
    try {
      const id = uniqueId || `${fieldName}_Error`;
      SetNotification(executionContext, fieldName, message, id);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Field.ShowError");
    }
  }

  export function ShowInfo(
    executionContext: Xrm.Events.EventContext,
    fieldName: string,
    message: string,
    uniqueId?: string
  ): void {
    try {
      const id = uniqueId || `${fieldName}_Info`;
      SetNotification(executionContext, fieldName, message, id);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Field.ShowInfo");
    }
  }

  export function Clear(
    executionContext: Xrm.Events.EventContext,
    fieldName: string,
    uniqueId?: string
  ): void {
    try {
      ClearNotification(executionContext, fieldName, uniqueId);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Field.Clear");
    }
  }
}



}
//#endregion

//#region Tabs
export namespace Tabs {

  export function HideTabs(executionContext: Xrm.Events.EventContext, ...tabNames: string[]): void {
    if (!executionContext || tabNames.length === 0) {
      Notifications.Form.ShowError(executionContext, "Execution context or tab names missing.", "Tabs");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (Generic.IsNull(formContext, "Form context is null.")) return;
      Internal.ShowHideTabs(formContext, false, ...tabNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Tabs.HideTabs");
    }
  }

  export function ShowTabs(executionContext: Xrm.Events.EventContext, ...tabNames: string[]): void {
    if (!executionContext || tabNames.length === 0) {
      Notifications.Form.ShowError(executionContext, "Execution context or tab names missing.", "Tabs");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (Generic.IsNull(formContext, "Form context is null.")) return;
      Internal.ShowHideTabs(formContext, true, ...tabNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Tabs.ShowTabs");
    }
  }

  export function SetFocus(executionContext: Xrm.Events.EventContext, tabName: string): void {
    if (!executionContext || Generic.IsNull(tabName)) {
      Notifications.Form.ShowError(executionContext, "Execution context or tab name missing.", "Tabs");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetTabFocus(formContext, tabName);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Tabs.SetFocus");
    }
  }

  export function SetLabel(executionContext: Xrm.Events.EventContext, tabName: string, label: string): void {
    if (!executionContext || Generic.IsNull(tabName) || Generic.IsNull(label)) {
      Notifications.Form.ShowError(executionContext, "Execution context, tab name, or label missing.", "Tabs");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.SetTabLabel(formContext, tabName, label);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Tabs.SetLabel");
    }
  }

}
//#endregion 

//#region Form
export namespace Form {

export function save(executionContext: Xrm.Events.EventContext): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");

      // Use explicit SaveOptions to satisfy typings
      const saveOptions: Xrm.SaveOptions = { saveMode: 1 }; // 1 = Save
      formContext.data.save(saveOptions);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.save");
    }
  }
  
  export async function saveAndClose(executionContext: Xrm.Events.EventContext): Promise<void> {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");

      const saveOptions: Xrm.SaveOptions = { saveMode: 1 };
      await formContext.data.save(saveOptions);
      formContext.ui.close();
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.saveAndClose");
    }
  }

  export async function refresh(executionContext: Xrm.Events.EventContext): Promise<void> {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");

      await formContext.data.refresh(false);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.refresh");
    }
  }

  export function getFormType(executionContext: Xrm.Events.EventContext): number | null {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      return formContext.ui.getFormType();
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.getFormType");
      return null;
    }
  }

  export function refreshSubGrid(executionContext: Xrm.Events.EventContext, ...gridNames: string[]): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      Internal.RefreshSubgrids(formContext, ...gridNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.refreshSubGrid");
    }
  }

  export function getSubGridRows(executionContext: Xrm.Events.EventContext, gridName: string): any | null {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      const gridContext = Internal.GetGridContext(formContext, gridName);
      if (!gridContext) throw new Error(`Grid '${gridName}' not found.`);
      return gridContext.getGrid().getRows();
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.getSubGridRows");
      return null;
    }
  }

  export function getSelectedSubGridRows(executionContext: Xrm.Events.EventContext, gridName: string): any | null {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      const gridContext = Internal.GetGridContext(formContext, gridName);
      if (!gridContext) throw new Error(`Grid '${gridName}' not found.`);
      return gridContext.getGrid().getSelectedRows();
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.getSelectedSubGridRows");
      return null;
    }
  }

  export function preventSave(executionContext: Xrm.Events.EventContext): void {
    try {
      const ctx = executionContext as any;
      const eventArgs = typeof ctx.getEventArgs === "function" ? ctx.getEventArgs() : undefined;
      if (eventArgs && typeof eventArgs.preventDefault === "function") {
        eventArgs.preventDefault();
      } else {
        throw new Error("Event arguments not available — ensure this is called from onSave event.");
      }
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.preventSave");
    }
  }

  export function ribbonRefresh(executionContext: Xrm.Events.EventContext): void {
    try {
      const formContext = Internal.GetFormContext(executionContext);
      if (!formContext) throw new Error("Form context not found.");
      formContext.ui.refreshRibbon();
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Form.ribbonRefresh");
    }
  }

}
//#endregion

//#region quick form
export namespace QuickForms {
  
  export function HideQuickForms(executionContext: Xrm.Events.EventContext, ...quickFormNames: string[]): void {
    if (!executionContext || quickFormNames.length === 0) {
      Generic.HandleError(executionContext, new Error("Execution context or quickForm names missing."), "QuickForms.HideQuickForms");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideQuickForms(formContext, false, ...quickFormNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "QuickForms.HideQuickForms");
    }
  }

  export function ShowQuickForms(executionContext: Xrm.Events.EventContext, ...quickFormNames: string[]): void {
    if (!executionContext || quickFormNames.length === 0) {
      Generic.HandleError(executionContext, new Error("Execution context or quickForm names missing."), "QuickForms.ShowQuickForms");
      return;
    }

    try {
      const formContext = Internal.GetFormContext(executionContext);
      Internal.ShowHideQuickForms(formContext, true, ...quickFormNames);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "QuickForms.ShowQuickForms");
    }
  }

}
//#endregion

//#region Alerts
export namespace Alerts {


  export async function OpenAlertDialog(
    executionContext: Xrm.Events.EventContext,
    message: string,
    title: string = "Alert",
    confirmLabel: string = "OK",
    height?: number,
    width?: number
  ): Promise<void> {
    try {
      if (Generic.IsNull(message, "Alert message is null.")) return;

      const alertStrings: Xrm.Navigation.AlertStrings = {
        text: message,
        title
      };

      const alertOptions: any = {
        confirmButtonLabel: confirmLabel
      };

      if (height !== undefined) alertOptions.height = height;
      if (width !== undefined) alertOptions.width = width;

      await Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Alerts.OpenAlertDialog");
    }

    // Always return explicitly (to satisfy TS return rules)
    return;
  }

 
  export async function OpenConfirmDialog(
    executionContext: Xrm.Events.EventContext,
    message: string,
    title: string = "Confirm Action",
    confirmLabel: string = "Yes",
    cancelLabel: string = "No",
    height?: number,
    width?: number
  ): Promise<{ confirmed: boolean }> {
    try {
      if (Generic.IsNull(message, "Confirm message is null.")) {
        return { confirmed: false };
      }

      const confirmStrings: Xrm.Navigation.ConfirmStrings = {
        text: message,
        title
      };

      const confirmOptions: any = {
        confirmButtonLabel: confirmLabel,
        cancelButtonLabel: cancelLabel
      };

      if (height !== undefined) confirmOptions.height = height;
      if (width !== undefined) confirmOptions.width = width;

      const result = await Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions);
      return (result as any) ?? { confirmed: false };
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Alerts.OpenConfirmDialog");
      return { confirmed: false };
    }
  }

}
//#endregion

//#region Data (crud operations)
export namespace Data {

  // CREATE
  export async function Create(
    executionContext: Xrm.Events.EventContext,
    entityLogicalName: string,
    data: any
  ): Promise<{ id: string } | null> {
    try {
      if (!entityLogicalName) throw new Error("entityLogicalName is required.");
      if (!data) throw new Error("data is required.");

      const api = Internal.GetWebAPIContext();
      const response = await api.createRecord(entityLogicalName, data);
      return response;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.Create");
      return null;
    }
  }

  // RETRIEVE
  export async function Retrieve(
    executionContext: Xrm.Events.EventContext,
    entityLogicalName: string,
    id: string,
    query?: string
  ): Promise<any | null> {
    try {
      if (!entityLogicalName) throw new Error("entityLogicalName is required.");
      if (!id) throw new Error("id is required.");

      const api = Internal.GetWebAPIContext();
      const response = await api.retrieveRecord(entityLogicalName, id, query);
      return response;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.Retrieve");
      return null;
    }
  }

  // UPDATE
  export async function Update(
    executionContext: Xrm.Events.EventContext,
    entityLogicalName: string,
    id: string,
    data: any
  ): Promise<void> {
    try {
      if (!entityLogicalName) throw new Error("entityLogicalName is required.");
      if (!id) throw new Error("id is required.");
      if (!data) throw new Error("data is required.");

      const api = Internal.GetWebAPIContext();
      await api.updateRecord(entityLogicalName, id, data);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.Update");
    }
  }

  // DELETE
  export async function Delete(
    executionContext: Xrm.Events.EventContext,
    entityLogicalName: string,
    id: string
  ): Promise<void> {
    try {
      if (!entityLogicalName) throw new Error("entityLogicalName is required.");
      if (!id) throw new Error("id is required.");

      const api = Internal.GetWebAPIContext();
      await api.deleteRecord(entityLogicalName, id);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.Delete");
    }
  }

  // RETRIEVE MULTIPLE
  export async function RetrieveMultiple(
    executionContext: Xrm.Events.EventContext,
    entityLogicalName: string,
    query?: string
  ): Promise<{ entities: any[] } | null> {
    try {
      if (!entityLogicalName) throw new Error("entityLogicalName is required.");

      const api = Internal.GetWebAPIContext();
      const response = await api.retrieveMultipleRecords(entityLogicalName, query || "");
      return response;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.RetrieveMultiple");
      return null;
    }
  }

  // EXECUTE SINGLE REQUEST
  export async function Execute(
    executionContext: Xrm.Events.EventContext,
    request: any
  ): Promise<Xrm.ExecuteResponse | null> {
    try {
      if (!request) throw new Error("request is required.");

      const api = Internal.GetWebAPIContext();
      const response = await (api as Xrm.WebApiOnline).execute(request);
      return response as Xrm.ExecuteResponse;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.Execute");
      return null;
    }
  }

  // EXECUTE MULTIPLE REQUESTS
  export async function ExecuteMultiple(
    executionContext: Xrm.Events.EventContext,
    requests: any[]
  ): Promise<Xrm.ExecuteResponse[]> {
    try {
      if (!Array.isArray(requests) || requests.length === 0) {
        throw new Error("requests must be a non-empty array.");
      }

      const api = Internal.GetWebAPIContext();
      const executions = requests.map(req =>
        (api as Xrm.WebApiOnline).execute(req)
      );

      return await Promise.all(executions);
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.ExecuteMultiple");
      return [];
    }
  }

  // EXECUTE GLOBAL ACTION REQUEST
  export async function ExecuteGlobalActionRequest(
    executionContext: Xrm.Events.EventContext,
    actionName: string,
    data: any
  ): Promise<Xrm.ExecuteResponse | null> {
    try {
      if (!actionName) throw new Error("actionName is required.");

      const request = {
        Input: data,
        getMetadata: () => ({
          boundParameter: null,
          parameterTypes: {
            Input: { typeName: "Edm.String", structuralProperty: 1 }
          },
          operationType: 0,
          operationName: actionName
        })
      };

      const api = Internal.GetWebAPIContext();
      const response = await (api as Xrm.WebApiOnline).execute(request);
      return response as Xrm.ExecuteResponse;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.ExecuteGlobalActionRequest");
      return null;
    }
  }

  // EXECUTE ENTITY ACTION REQUEST
  export async function ExecuteEntityActionRequest(
    executionContext: Xrm.Events.EventContext,
    actionName: string,
    entityName: string,
    id: string,
    data: any
  ): Promise<Xrm.ExecuteResponse | null> {
    try {
      if (!actionName) throw new Error("actionName is required.");
      if (!entityName) throw new Error("entityName is required.");
      if (!id) throw new Error("id is required.");

      const request = {
        entity: { id, entityType: entityName },
        Input: data,
        getMetadata: () => ({
          boundParameter: "entity",
          parameterTypes: {
            entity: { typeName: `mscrm.${entityName}`, structuralProperty: 5 },
            Input: { typeName: "Edm.String", structuralProperty: 1 }
          },
          operationType: 0,
          operationName: actionName
        })
      };

      const api = Internal.GetWebAPIContext();
      const response = await (api as Xrm.WebApiOnline).execute(request);
      return response as Xrm.ExecuteResponse;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "Data.ExecuteEntityActionRequest");
      return null;
    }
  }
}
//#endregion

//#region UserInfo

export namespace UserInfo {

  /** Checks if current user has a specific role */
  export function userHasRole(executionContext: Xrm.Events.EventContext, roleName: string): boolean {
    try {
      if (!roleName) throw new Error("Role name must be provided.");

      const userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

      if (!userRoles || userRoles.getLength() === 0) return false;

      for (let i = 0; i < userRoles.getLength(); i++) {
        const role = userRoles.get(i);
        const roleNameInSystem = (role as any).name;

        if (roleNameInSystem && roleNameInSystem.toLowerCase() === roleName.toLowerCase()) {
          return true;
        }
      }

      return false;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "UserInfo.userHasRole");
      return false;
    }
  }

  /** Checks if user has any one of the given roles */
  export function userHasAnyRole(executionContext: Xrm.Events.EventContext, ...roleNames: string[]): boolean {
    try {
      if (roleNames.length === 0) return false;

      const userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
      if (!userRoles || userRoles.getLength() === 0) return false;

      const lowerRoleNames = roleNames.map(r => r.toLowerCase());

      for (let i = 0; i < userRoles.getLength(); i++) {
        const role = userRoles.get(i);
        const roleNameInSystem = (role as any).name;

        if (roleNameInSystem && lowerRoleNames.indexOf(roleNameInSystem.toLowerCase()) !== -1) {
          return true;
        }
      }

      return false;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "UserInfo.userHasAnyRole");
      return false;
    }
  }

  /** Gets current user ID (without braces) */
  export function getUserId(executionContext: Xrm.Events.EventContext): string | null {
    try {
      const id = Xrm.Utility.getGlobalContext().userSettings.userId;
      return id ? id.replace(/[{}]/g, "") : null;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "UserInfo.getUserId");
      return null;
    }
  }

  /** Gets current user full name */
  export function getUserName(executionContext: Xrm.Events.EventContext): string | null {
    try {
      const name = Xrm.Utility.getGlobalContext().userSettings.userName;
      return name || null;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "UserInfo.getUserName");
      return null;
    }
  }

  /** Returns list of all role names assigned to current user */
  export function getUserRoles(executionContext: Xrm.Events.EventContext): string[] {
    try {
      const userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
      const roleNames: string[] = [];

      if (!userRoles || userRoles.getLength() === 0) return roleNames;

      for (let i = 0; i < userRoles.getLength(); i++) {
        const role = userRoles.get(i);
        const name = (role as any).name;
        if (name) roleNames.push(name);
      }

      return roleNames;
    } catch (error: any) {
      Generic.HandleError(executionContext, error, "UserInfo.getUserRoles");
      return [];
    }
  }
}
//#endregion

//#region Security Roles
export namespace Security {
  export namespace Roles {
    export const Agent: string = "jetAgent";
    export const Manager: string = "jetManager";
    export const Admin: string = "jetAdmin";
  }
}
//#endregion 

}
(window as any).CommonD365 = CommonD365;

