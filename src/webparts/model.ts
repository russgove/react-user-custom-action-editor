import { IUserCustomAction, IUserCustomActionInfo, UserCustomActionRegistrationType, UserCustomActionScope } from '@pnp/sp/user-custom-actions';

export interface IActionRef {
    ActionInfo:IUserCustomActionInfo;
    Source: "site"|"web"|"list";
    SourcId: string;
}