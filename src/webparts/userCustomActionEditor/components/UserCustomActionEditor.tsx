import * as React from 'react';
import styles from './UserCustomActionEditor.module.scss';
import { IUserCustomActionEditorProps } from './IUserCustomActionEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";

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
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Overlay } from "office-ui-fabric-react/lib/Overlay";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
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
        debugger; setCommand("add"); setSelectedUserCustomAction({
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
  const farItemsNonFocusable: IContextualMenuItem[] = [
    {
      key: "Update Selected",
      name: "Update Selected",
      icon: "TriggerApproval",
      //  disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes",
      onClick: (e) => {

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
      key: "edit", name: "Edit",
      isResizable: false,
      fieldName: "dummy", minWidth: 70,
      onRender: (item?: any, index?: number, column?: IColumn) => {
        return (
          <IconButton iconProps={{ iconName: "Edit" }} onClick={(e) => { setSelectedUserCustomAction(item); setCommand("edit"); }} />
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
  debugger;

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
      <Panel
        type={PanelType.custom | PanelType.smallFixedNear}
        customWidth='900px'
        isOpen={command !== null}
        headerText={command === "add" ? "Add new Action" : "Edit existing Action"}
        onDismiss={
          () => {
            setSelectedUserCustomAction(null); setCommand(null);
          }}
        isBlocking={true}
      >
          <TextField label='Id' value={selectedUserCustomAction ? selectedUserCustomAction["Id"] : ""} />
      
        <TextField label='ClientSideComponentId' value={selectedUserCustomAction ? selectedUserCustomAction["ClientSideComponentId"] : ""} />
        <TextField label='ClientSideComponentProperties' multiline={true} value={selectedUserCustomAction ? selectedUserCustomAction["ClientSideComponentProperties"] : ""} />
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
        }}  />
        <TextField label='HostProperties' value={selectedUserCustomAction ? selectedUserCustomAction["HostProperties"] : ""} />
        <TextField label='ImageUrl' value={selectedUserCustomAction ? selectedUserCustomAction["ImageUrl"] : ""} />
        <TextField label='Location' value={selectedUserCustomAction ? selectedUserCustomAction["Location"] : ""} 
             onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({ ...selectedUserCustomAction, Location: newValue });
            }} 
        />
        <TextField label='Name' value={selectedUserCustomAction ? selectedUserCustomAction["Name"] : ""} />
        <TextField label='RegistrationId' value={selectedUserCustomAction ? selectedUserCustomAction["RegistrationId"] : ""} />
        <TextField label='RegistrationType' value={selectedUserCustomAction ? selectedUserCustomAction["RegistrationType"].toString() : ""} />
        <TextField label='Scope' value={selectedUserCustomAction ? selectedUserCustomAction["Scope"].toString() : ""} />
        <TextField label='ScriptBlock' value={selectedUserCustomAction ? selectedUserCustomAction["ScriptBlock"] : ""} />
        <TextField label='ScriptSrc' value={selectedUserCustomAction ? selectedUserCustomAction["ScriptSrc"] : ""} />
        <TextField label='Sequence' value={selectedUserCustomAction ? selectedUserCustomAction["Sequence"].toString() : ""} />
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
          debugger
          const newValues = {
            "CommandUIExtension": selectedUserCustomAction.CommandUIExtension,
            "Description": selectedUserCustomAction.Description,
            "Group": selectedUserCustomAction.Group,
            "ImageUrl": selectedUserCustomAction.ImageUrl,
            "Location": selectedUserCustomAction.Location,
            "Name": selectedUserCustomAction.Name,
            "RegistrationId": selectedUserCustomAction.RegistrationId,
            "RegistrationType": selectedUserCustomAction.RegistrationType,
            "Rights": selectedUserCustomAction.Rights,
            "Scope": selectedUserCustomAction.Scope,
            "ScriptBlock": selectedUserCustomAction.ScriptBlock,
            "ScriptSrc": selectedUserCustomAction.ScriptSrc,
            "Sequence": selectedUserCustomAction.Sequence,
            "Title": selectedUserCustomAction.Title,
            "Url": selectedUserCustomAction.Url,
            "VersionOfUserCustomAction": selectedUserCustomAction.VersionOfUserCustomAction

          };
          if (command === 'edit') {
            newValues["Id"]= selectedUserCustomAction.Id;
          
            await spWeb.web.userCustomActions.getById(selectedUserCustomAction.Id).update(newValues)
              .then((x) => { debugger; })
              .catch((x) => { debugger; })
          } else  if (command === 'add') {
           await spWeb.web.userCustomActions.add(newValues)
              .then((x) => { debugger; })
              .catch((x) => { debugger; })
          }
          setRefresh(!refresh);
        }
        } >Save</PrimaryButton>
      </Panel>
    </div>
  );
}

