import * as React from 'react';
import styles from './UserCustomActionEditor.module.scss';
import { IUserCustomActionEditorProps } from './IUserCustomActionEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { forEach, map, forIn } from 'lodash';
import { spfi, SPFx } from "@pnp/sp";
import { SPPermission } from "@microsoft/sp-page-context";
import { PermissionKind, IBasePermissions } from "@pnp/sp/security";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/user-custom-actions";
import { IUserCustomAction, IUserCustomActionInfo, UserCustomActionRegistrationType, UserCustomActionScope } from '@pnp/sp/user-custom-actions';
import { useState, useEffect } from "react";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import { DetailsList, IColumn, Selection, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Dialog, DialogFooter, DialogType } from "office-ui-fabric-react/lib/Dialog";
import { ScrollablePane } from "office-ui-fabric-react/lib/ScrollablePane";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Overlay } from "office-ui-fabric-react/lib/Overlay";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { ComboBox, Dropdown, IComboBoxOption } from 'office-ui-fabric-react';
export default function UserCustomActionEditor(props: IUserCustomActionEditorProps) {

  const [command, setCommand] = React.useState<string>(null);
  const [sortBy, setSortBy] = React.useState<string>("Title");

  const [sortDescending, setSortDescending] = React.useState<boolean>(false);
  const [refresh, setRefresh] = React.useState<boolean>(false);
  const [selectedUserCustomAction, setSelectedUserCustomAction] = React.useState<IUserCustomActionInfo>(null);
  const itemsNonFocusable: IContextualMenuItem[] = [
    {
      key: "Add New",
      name: "Add New",
      icon: "Edit",
      // disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes",
      onClick: (e) => {
        setCommand("add"); setSelectedUserCustomAction({
          CommandUIExtension: null,
          Description: null,
          Group: null,
          Id: null,
          ImageUrl: null,
          Location: null,
          Name: null,
          RegistrationId: null,
          RegistrationType: 0,
          Rights: null,
          Scope: 0,
          ScriptBlock: null,
          ScriptSrc: null,
          Sequence: 0,
          Title: null,
          Url: null,
          VersionOfUserCustomAction: null,
        });
      },

    }];
  /**
* Taken from sp.js, checks the supplied permissions against the mask
*
* @param value The security principal's permissions on the given object
* @param perm The permission checked against the value
*/
 

  function isSelectedPermission(ucaPermission: IBasePermissions, cbxPermission: SPPermission): boolean {
    // not checking if we have permission, just WHAT THE PERMISSION IS!!

    // full mask is not a full mask (dang bit-shifters!)
    if (cbxPermission.value.High === 2147483647 && cbxPermission.value.Low === 4294967295) {
      if (ucaPermission.Low as unknown == '65535' && ucaPermission.High as unknown == '32767')
        return true;
   }
    if (
      ((ucaPermission.Low) == (cbxPermission.value.Low))
      &&
      ((ucaPermission.High) == (cbxPermission.value.High))
    ) {
      return true;
    }

    return false;
  }
  function getPermissionKindsOfUCA(uca: IUserCustomActionInfo): string[] {
    const whatitgot: Array<string> = [];
  
    Object.keys(SPPermission).forEach(key => {
      console.log(key);
      console.log(SPPermission[key]);
      if (isSelectedPermission(uca.Rights, SPPermission[key])) {
        whatitgot.push(key);
      }
    });
    return whatitgot;
  }

  const spPermissions: IComboBoxOption[] = Object.keys(SPPermission).map(key => {
    return ({ key: key, text: key, data: SPPermission[key] })
  });



  function getHeader(command: string): string {
    switch (command) {
      case "add": { return "Add new Action" }
      case "edit": { return "Edit existing Action" }
      case "delete": { return "Delete Action" }
      default: { return "unknown  command" }
    }
  }
  function getButtonText(command: string): string {
    switch (command) {
      case "add": { return "Add Action" }
      case "edit": { return "Save changes" }
      case "delete": { return "Delete Action" }
      default: { return "unknown  command" }
    }
  }

  const farItemsNonFocusable: IContextualMenuItem[] = [
    {
      key: "Refresh",
      name: "Refresh",
      icon: "Refresh",
      //  disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes",
      onClick: (e) => {
        setRefresh(!refresh)
      },

    }];
  function columnHeaderClicked(ev: React.MouseEvent<HTMLElement>, column: IColumn) {
    if (sortBy === column.fieldName) {
      setSortDescending(!sortDescending);
    }
    else {
      setSortBy(column.fieldName);
      setSortDescending(false);
    }
  }
  const cols: Array<IColumn> = [
    {
      key: "edit", name: "Actions",
      isResizable: false,
      fieldName: "dummy", minWidth: 70,
      onRender: (item?: any, index?: number, column?: IColumn) => {
        return (
          <div>
            <IconButton iconProps={{ iconName: "Edit" }} onClick={(e) => { setSelectedUserCustomAction(item); setCommand("edit"); }} />
            <IconButton iconProps={{ iconName: "Delete" }} onClick={(e) => { setSelectedUserCustomAction(item); setCommand("delete"); }} />
          </div>
        );

      }
    },
    {
      key: "Id", name: "Id", fieldName: "Id", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Id", onColumnClick: columnHeaderClicked,
    },
    {
      key: "Title", name: "Title", fieldName: "Title", minWidth: 200,
      isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Title", onColumnClick: columnHeaderClicked,
    },
    {
      key: "Location", name: "Location", fieldName: "Location", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Location", onColumnClick: columnHeaderClicked,
    }
  ];


  const [actions, setActions] = React.useState<Array<IUserCustomActionInfo>>([]);
  useEffect(
    () => {
      const spWeb = spfi().using(SPFx(props.context));
      spWeb.web.userCustomActions().then((ucas) => {
        setActions(ucas);
      })

    }, [refresh]);


  return (
    <div className={styles.userCustomActionEditor}>
      <CommandBar
        // isSearchBoxVisible={false}
        items={itemsNonFocusable}
        farItems={farItemsNonFocusable}

      />
      <DetailsList items={actions} columns={cols}></DetailsList>
      {(selectedUserCustomAction) &&
        <Panel
          type={PanelType.custom | PanelType.smallFixedNear}
          customWidth='900px'
          isOpen={command !== null}
          headerText={
            getHeader(command)

          }
          onDismiss={
            () => {
              setSelectedUserCustomAction(null); setCommand(null);
            }}
          isBlocking={true}
        >

          <TextField label='Id' disabled={true} value={selectedUserCustomAction ? selectedUserCustomAction["Id"] : ""} />
          <ComboBox label='Rights'
            options={spPermissions}
            multiSelect={true}
            selectedKey={getPermissionKindsOfUCA(selectedUserCustomAction)}
            onChange={(event, selection, xxx) => {
              debugger;
              var rights: any = {};
              if (selection.key === 'fullMask') {
                rights.High = '32767';
                rights.Low = '65535';
              }
              else {
                rights.High = selection.data.value.High.toString();
                rights.Low = selection.data.value.Low.toString()
              }
              console.log(rights);
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Rights: rights });
            }}
          />
          <TextField label='ClientSideComponentId' value={selectedUserCustomAction ? selectedUserCustomAction["ClientSideComponentId"] : ""} />
          <TextField label='ClientSideComponentProperties' multiline={true} value={selectedUserCustomAction ? selectedUserCustomAction["ClientSideComponentProperties"] : ""}
          />
          <TextField label='CommandUIExtension' multiline={true} value={selectedUserCustomAction ? selectedUserCustomAction["CommandUIExtension"] : ""} />
          <TextField label='Description'
            value={selectedUserCustomAction ? selectedUserCustomAction["Description"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Description: newValue });
            }}
          />
          <TextField label='Group' value={selectedUserCustomAction ? selectedUserCustomAction["Group"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Group: newValue });
            }} />
          <TextField label='HostProperties' value={selectedUserCustomAction ? selectedUserCustomAction["HostProperties"] : ""} />
          <TextField label='ImageUrl' value={selectedUserCustomAction ? selectedUserCustomAction["ImageUrl"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, ImageUrl: newValue });
            }}
          />
          <TextField label='Location' value={selectedUserCustomAction ? selectedUserCustomAction["Location"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Location: newValue });
            }}
          />
          <TextField label='Name' value={selectedUserCustomAction ? selectedUserCustomAction["Name"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Name: newValue });
            }}
          />
          <TextField label='RegistrationId' value={selectedUserCustomAction ? selectedUserCustomAction["RegistrationId"] : ""} />
          <TextField label='RegistrationType' value={selectedUserCustomAction ? selectedUserCustomAction["RegistrationType"].toString() : ""} />
          {/* <TextField label='Rights' value={selectedUserCustomAction ? selectedUserCustomAction["Rights"].toString() : ""} /> */}




          <TextField label='Scope' value={selectedUserCustomAction ? selectedUserCustomAction["Scope"].toString() : ""} />
          <TextField label='ScriptBlock' value={selectedUserCustomAction ? selectedUserCustomAction["ScriptBlock"] : ""} />
          <TextField label='ScriptSrc' value={selectedUserCustomAction ? selectedUserCustomAction["ScriptSrc"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, ScriptSrc: newValue });
            }} />
          <TextField label='Sequence' type='number' value={selectedUserCustomAction ? selectedUserCustomAction["Sequence"].toString() : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Sequence: parseInt(newValue) });
            }}
          />
          <TextField label='Title' value={selectedUserCustomAction ? selectedUserCustomAction["Title"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Title: newValue });
            }}
          />
          <TextField label='Url' value={selectedUserCustomAction ? selectedUserCustomAction["Url"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Url: newValue });
            }} />
          <TextField label='VersionOfUserCustomAction' value={selectedUserCustomAction ? selectedUserCustomAction["VersionOfUserCustomAction"] : ""} />

          <PrimaryButton onClick={async (e) => {
            const spWeb = spfi().using(SPFx(props.context));

            const newValues = {};
            if (selectedUserCustomAction.CommandUIExtension) newValues["CommandUIExtension"] = selectedUserCustomAction.CommandUIExtension;
            if (selectedUserCustomAction.Description) newValues["Description"] = selectedUserCustomAction.Description;
            if (selectedUserCustomAction.Group) newValues["Group"] = selectedUserCustomAction.Group;
            if (selectedUserCustomAction.ImageUrl) newValues["ImageUrl"] = selectedUserCustomAction.ImageUrl;
            if (selectedUserCustomAction.Location) newValues["Location"] = selectedUserCustomAction.Location;
            if (selectedUserCustomAction.Name) newValues["Name"] = selectedUserCustomAction.Name;
            if (selectedUserCustomAction.RegistrationId) newValues["RegistrationId"] = selectedUserCustomAction.RegistrationId;
            if (selectedUserCustomAction.RegistrationType) newValues["RegistrationType"] = selectedUserCustomAction.RegistrationType;
            if (selectedUserCustomAction.Rights) newValues["Rights"] = selectedUserCustomAction.Rights;
            if (selectedUserCustomAction.Scope) newValues["Scope"] = selectedUserCustomAction.Scope;
            if (selectedUserCustomAction.ScriptBlock) newValues["ScriptBlock"] = selectedUserCustomAction.ScriptBlock;
            if (selectedUserCustomAction.ScriptSrc) newValues["ScriptSrc"] = selectedUserCustomAction.ScriptSrc;
            if (selectedUserCustomAction.Sequence) newValues["Sequence"] = selectedUserCustomAction.Sequence;
            if (selectedUserCustomAction.Title) newValues["Title"] = selectedUserCustomAction.Title;
            if (selectedUserCustomAction.Url) newValues["Url"] = selectedUserCustomAction.Url;
            if (selectedUserCustomAction.VersionOfUserCustomAction) newValues["VersionOfUserCustomAction"] = selectedUserCustomAction.VersionOfUserCustomAction


            switch (command) {
              case 'edit': {
                newValues["Id"] = selectedUserCustomAction.Id;
                await spWeb.web.userCustomActions.getById(selectedUserCustomAction.Id).update(newValues)
                  .catch((x) => { debugger; });
                break;
              }
              case 'add': {
                await spWeb.web.userCustomActions.add(newValues)
                  .catch((x) => { debugger; });
                break;
              }
              case 'delete': {
                newValues["Id"] = selectedUserCustomAction.Id;
                await spWeb.web.userCustomActions.getById(selectedUserCustomAction.Id).delete()
                  .catch((x) => { debugger; });
                break;
              }
            }
            setRefresh(!refresh); setCommand(null);
          }
          } >{getButtonText(command)}</PrimaryButton>

        </Panel>
      }
    </div >

  );
}

