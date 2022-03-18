import { IUserCustomAction, IUserCustomActionInfo, UserCustomActionRegistrationType, UserCustomActionScope } from '@pnp/sp/user-custom-actions';

export interface IUserCustomActionReference {
    CustomActuion:IUserCustomActionInfo;
    Source: "site"|"web"|"list";
    SourcId: string;
}