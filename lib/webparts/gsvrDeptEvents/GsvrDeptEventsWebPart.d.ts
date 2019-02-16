import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface ISPLists {
    value: ISPList[];
}
export interface ISPList {
    Title: string;
    Id: string;
    AnncURL: string;
    DeptURL: string;
    CalURL: string;
    a85u: string;
}
export interface IGsvrDeptEventsWebPartProps {
    description: string;
}
export default class GsvrDeptEventsWebPart extends BaseClientSideWebPart<IGsvrDeptEventsWebPartProps> {
    getuser: Promise<{}>;
    render(): void;
    protected readonly dataVersion: Version;
    _getListData(): Promise<ISPLists>;
    private _renderList(items);
    onInit(): Promise<void>;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
