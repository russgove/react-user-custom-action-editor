import * as React from 'react';
import styles from './UserCustomActionEditor.module.scss';
import { IUserCustomActionEditorProps } from './IUserCustomActionEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";

import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/user-custom-actions";
import { IUserCustomAction, IUserCustomActionInfo } from '@pnp/sp/user-custom-actions';
import { useState, useEffect } from "react";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import { DetailsList, IColumn, Selection, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Dialog, DialogFooter, DialogType } from "office-ui-fabric-react/lib/Dialog";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Overlay } from "office-ui-fabric-react/lib/Overlay";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
export default function UserCustomActionEditor(props: IUserCustomActionEditorProps) {
  const [sortBy, setSortBy] = React.useState<string>("Title");
  const [sortDescending, setSortDescending] = React.useState<boolean>(false);
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
    { key: "Id", name: "Id", fieldName: "Id", minWidth: 200,    isSortedDescending: sortDescending,
    sortAscendingAriaLabel: 'Sorted A to Z',
    sortDescendingAriaLabel: 'Sorted Z to A',
    isSorted: sortBy === "Id", onColumnClick: columnHeaderClicked, },
    {
      key: "Title", name: "Title", fieldName: "Title", minWidth: 200,
      isSortedDescending: sortDescending,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      isSorted: sortBy === "Title", onColumnClick: columnHeaderClicked,
    },
    { key: "Location", name: "Location", fieldName: "Location", minWidth: 200,    isSortedDescending: sortDescending,
    sortAscendingAriaLabel: 'Sorted A to Z',
    sortDescendingAriaLabel: 'Sorted Z to A',
    isSorted: sortBy === "Location", onColumnClick: columnHeaderClicked, }
  ];
  debugger;

  const [actions, setActions] = React.useState<Array<IUserCustomActionInfo>>([]);
  useEffect(
    () => {
      const spWeb = spfi().using(SPFx(props.context));
      spWeb.web.userCustomActions().then((ucas) => {
        setActions(ucas);
      })

    }, []);
  return (
    <div className={styles.userCustomActionEditor}>
      <DetailsList items={actions} columns={cols}></DetailsList>
    </div>
  );
}

