import * as React from 'react';
import styles from './UserCustomActionEditor.module.scss';
import { IUserCustomActionEditorProps } from './IUserCustomActionEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { forEach, map, forIn } from 'lodash';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { SPPermission } from "@microsoft/sp-page-context";
import { PermissionKind, IBasePermissions } from "@pnp/sp/security";
import { IWebInfo, Web } from "@pnp/sp/webs";
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import "@pnp/sp/webs";
import "@pnp/sp/user-custom-actions";
import { IUserCustomAction, IUserCustomActionInfo, UserCustomActionRegistrationType, UserCustomActionScope } from '@pnp/sp/user-custom-actions';
import { useState, useEffect } from "react";
import { IActionRef } from "../../model";
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
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/batching";
import "@pnp/sp/sites";
import { IContextInfo } from "@pnp/sp/sites";
export default function UserCustomActionEditor(props: IUserCustomActionEditorProps) {

  // Helper https://stackoverflow.com/questions/43100718/typescript-enum-to-object-array
  const StringIsNumber = value => isNaN(Number(value)) === false;

  // Turn enum into array
  function ToArray(enumme) {
    return Object.keys(enumme)
      .filter(StringIsNumber)
      .map(key => {
        return { key: key, value: enumme[key] }
      });
  }
  const userCustomActionScopes = ToArray(UserCustomActionScope);
  const userCustomActionRegistrationTypes = ToArray(UserCustomActionRegistrationType);
  debugger;

  const [webInfo, setWebInfo] = React.useState<IWebInfo>(null);
  const [command, setCommand] = React.useState<string>(null);
  const [sortBy, setSortBy] = React.useState<string>("Title");
  const [actions, setActions] = React.useState<Array<IActionRef>>([]);
  const [sortDescending, setSortDescending] = React.useState<boolean>(false);
  const [refresh, setRefresh] = React.useState<boolean>(false);
  const [selectedUserCustomAction, setSelectedUserCustomAction] = React.useState<IActionRef>(null);
  const itemsNonFocusable: IContextualMenuItem[] = [
    {
      key: "Add New",
      name: "Add New",
      icon: "Edit",
      // disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes",
      onClick: (e) => {
        setCommand("add"); setSelectedUserCustomAction({
          Source: null, SourcId: null, ActionInfo: {
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
          }
        });
      },

    }];

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
      key: "Id", name: "Id", fieldName: "CustomActuion.Id", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Id", onColumnClick: columnHeaderClicked,
      onRender: (item?: IActionRef, index?: number, column?: IColumn) => { return item.ActionInfo.Id }
    },
    {
      key: "Title", name: "Title", fieldName: "CustomActuion.Title", minWidth: 200,
      isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Title", onColumnClick: columnHeaderClicked,
      onRender: (item?: IActionRef, index?: number, column?: IColumn) => { return item.ActionInfo.Title }

    },
    {
      key: "Scope", name: "Scope", fieldName: "Scope", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Location", onColumnClick: columnHeaderClicked,
      onRender: (item?: IActionRef, index?: number, column?: IColumn) => { return UserCustomActionScope[item.ActionInfo.Scope] }

    },
    {
      key: "Scope", name: "Scope", fieldName: "Scope", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Location", onColumnClick: columnHeaderClicked,
      onRender: (item?: IActionRef, index?: number, column?: IColumn) => { return UserCustomActionRegistrationType[item.ActionInfo.RegistrationType] }
    },
    {
      key: "Source", name: "Source", fieldName: "Source", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Location", onColumnClick: columnHeaderClicked,
      onRender: (item?: IActionRef, index?: number, column?: IColumn) => { return item.Source }

    },
    {
      key: "Location", name: "Location", fieldName: "CustomAction.Location", minWidth: 200, isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Location", onColumnClick: columnHeaderClicked,
      onRender: (item?: IActionRef, index?: number, column?: IColumn) => { return item.ActionInfo.Location }

    }
  ];



  useEffect(() => {
    debugger;
    const fetchData = async () => {
      var spWeb: SPFI;
      const searchParams = new URLSearchParams(window.location.search);
      if (searchParams.get('site')) {
        spWeb = spfi(searchParams.get('site')).using(SPFx(props.context));

      } else {
        spWeb = spfi().using(SPFx(props.context));
      }
      const [batchedSP, execute] = spWeb.batched();


      await spWeb.web().then(r => {
        setWebInfo(r);
      })
        .catch((e) => {
          debugger;
        });
      batchedSP.web.userCustomActions().then((ucas) => {
        setActions(ucas.map(uca => { return { Source: 'web', ActionInfo: uca, SourcId: '' } }));
      });
      batchedSP.web.lists.expand("UserCustomActions")().then((listswithactions) => {
        debugger;
        //useeconst listActions:IActionRef=listswithactions.map((lwa=>{return{}}))
        console.log(listswithactions)
        listswithactions.map(uca => { return { Source: 'web', ActionInfo: uca, SourcId: webInfo.Id } });
      });
      execute();

    };

    fetchData();

  }, [refresh]);


  return (
    <div className={styles.userCustomActionEditor}>
      <h1>{webInfo ? webInfo.Title : "Loading..."}</h1>
      <h3>{webInfo ? webInfo.Url : "Loading..."}</h3>
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

          <TextField label='Id' disabled={true} value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Id"] : ""} />
          <ComboBox label='Rights'
            options={spPermissions}
            multiSelect={true}
            selectedKey={getPermissionKindsOfUCA(selectedUserCustomAction.ActionInfo)}
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
              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Rights: rights
                }
              });
            }}
          />
          <TextField label='ClientSideComponentId' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["ClientSideComponentId"] : ""} />
          <TextField label='ClientSideComponentProperties' multiline={true} value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["ClientSideComponentProperties"] : ""}
          />
          <TextField label='CommandUIExtension' multiline={true} value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["CommandUIExtension"] : ""} />
          <TextField label='Description'
            value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Description"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Description: newValue
                }
              });
            }}
          />
          <TextField label='Group' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Group"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Group: newValue
                }
              });
            }} />
          <TextField label='HostProperties' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["HostProperties"] : ""} />
          <TextField label='ImageUrl' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["ImageUrl"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  ImageUrl: newValue
                }
              });
            }}
          />
          <TextField label='Location' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Location"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Location: newValue
                }
              });
            }}
          />
          <TextField label='Name' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Name"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Name: newValue
                }
              });
            }}
          />
          <TextField label='RegistrationId' value={selectedUserCustomAction ? selectedUserCustomAction["RegistrationId"] : ""} />
          <TextField label='RegistrationType' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["RegistrationType"].toString() : ""} />
          {/* <TextField label='Rights' value={selectedUserCustomAction ? selectedUserCustomAction["Rights"].toString() : ""} /> */}




          <TextField label='Scope' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Scope"].toString() : ""} />
          <TextField label='ScriptBlock' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["ScriptBlock"] : ""} />
          <TextField label='ScriptSrc' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["ScriptSrc"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  ScriptSrc: newValue
                }
              });
            }} />
          <TextField label='Sequence' type='number' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Sequence"].toString() : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Sequence: parseInt(newValue)
                }
              });
            }}
          />
          <TextField label='Title' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Title"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Title: newValue
                }
              });
            }}
          />
          <TextField label='Url' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["Url"] : ""}
            onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setSelectedUserCustomAction({
                ...selectedUserCustomAction,
                ActionInfo: {
                  ...selectedUserCustomAction.ActionInfo,
                  Url: newValue
                }
              });
            }} />
          <TextField label='VersionOfUserCustomAction' value={selectedUserCustomAction ? selectedUserCustomAction.ActionInfo["VersionOfUserCustomAction"] : ""} />

          <PrimaryButton onClick={async (e) => {
            const spWeb = spfi().using(SPFx(props.context));

            const newValues = {};
            if (selectedUserCustomAction.ActionInfo.CommandUIExtension) newValues["CommandUIExtension"] = selectedUserCustomAction.ActionInfo.CommandUIExtension;
            if (selectedUserCustomAction.ActionInfo.Description) newValues["Description"] = selectedUserCustomAction.ActionInfo.Description;
            if (selectedUserCustomAction.ActionInfo.Group) newValues["Group"] = selectedUserCustomAction.ActionInfo.Group;
            if (selectedUserCustomAction.ActionInfo.ImageUrl) newValues["ImageUrl"] = selectedUserCustomAction.ActionInfo.ImageUrl;
            if (selectedUserCustomAction.ActionInfo.Location) newValues["Location"] = selectedUserCustomAction.ActionInfo.Location;
            if (selectedUserCustomAction.ActionInfo.Name) newValues["Name"] = selectedUserCustomAction.ActionInfo.Name;
            if (selectedUserCustomAction.ActionInfo.RegistrationId) newValues["RegistrationId"] = selectedUserCustomAction.ActionInfo.RegistrationId;
            if (selectedUserCustomAction.ActionInfo.RegistrationType) newValues["RegistrationType"] = selectedUserCustomAction.ActionInfo.RegistrationType;
            if (selectedUserCustomAction.ActionInfo.Rights) newValues["Rights"] = selectedUserCustomAction.ActionInfo.Rights;
            if (selectedUserCustomAction.ActionInfo.Scope) newValues["Scope"] = selectedUserCustomAction.ActionInfo.Scope;
            if (selectedUserCustomAction.ActionInfo.ScriptBlock) newValues["ScriptBlock"] = selectedUserCustomAction.ActionInfo.ScriptBlock;
            if (selectedUserCustomAction.ActionInfo.ScriptSrc) newValues["ScriptSrc"] = selectedUserCustomAction.ActionInfo.ScriptSrc;
            if (selectedUserCustomAction.ActionInfo.Sequence) newValues["Sequence"] = selectedUserCustomAction.ActionInfo.Sequence;
            if (selectedUserCustomAction.ActionInfo.Title) newValues["Title"] = selectedUserCustomAction.ActionInfo.Title;
            if (selectedUserCustomAction.ActionInfo.Url) newValues["Url"] = selectedUserCustomAction.ActionInfo.Url;
            if (selectedUserCustomAction.ActionInfo.VersionOfUserCustomAction) newValues["VersionOfUserCustomAction"] = selectedUserCustomAction.ActionInfo.VersionOfUserCustomAction


            switch (command) {
              case 'edit': {
                newValues["Id"] = selectedUserCustomAction.ActionInfo.Id;
                await spWeb.web.userCustomActions.getById(selectedUserCustomAction.ActionInfo.Id).update(newValues)
                  .catch((x) => { debugger; });
                break;
              }
              case 'add': {
                await spWeb.web.userCustomActions.add(newValues)
                  .catch((x) => { debugger; });
                break;
              }
              case 'delete': {
                newValues["Id"] = selectedUserCustomAction.ActionInfo.Id;
                await spWeb.web.userCustomActions.getById(selectedUserCustomAction.ActionInfo.Id).delete()
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

