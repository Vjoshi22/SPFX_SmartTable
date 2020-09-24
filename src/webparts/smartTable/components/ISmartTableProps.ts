import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISmartTableProps {
  description: string;
  tableTitle:string;
  context: WebPartContext;
  lists: string;
  columnData: any[];
  selectedKeys:any[];
  fieldListCollection: string[];
  color: string;
}
