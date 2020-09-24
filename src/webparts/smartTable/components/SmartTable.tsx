import * as React from 'react';
import styles from './SmartTable.module.scss';
import { ISmartTableProps } from './ISmartTableProps';
import { escape } from '@microsoft/sp-lodash-subset';
// other imports
import { ListPicker } from "@pnp/spfx-controls-react/lib/ListPicker";
//Import related to react-bootstrap-table-next    
import BootstrapTable from 'react-bootstrap-table-next';    
//Import from @pnp/sp    
import { sp } from "@pnp/sp";    
import "@pnp/sp/webs";    
import "@pnp/sp/lists/web";    
import "@pnp/sp/items/list";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { Web, List} from "sp-pnp-js";


export interface IUserItem{
  dataField:string;
  text:string
}
export interface IProjectMasterState{
  projectMasterList:any[]
  listguid: string;
}

var users: Array<IUserItem> = new Array<IUserItem>(); 
const projectMasterCol = [{
    dataField:'Temp', 
    text:'Temp'
}];
export default class SmartTable extends React.Component<ISmartTableProps, IProjectMasterState> {
  constructor(props: ISmartTableProps){    
    super(props);    
    this.state ={    
      projectMasterList : [],
      listguid: ''
    }    
  } 
  public componentWillMount() {
    this.getAllProjectDetails().then(resultList => {
      this.setState({    
        projectMasterList:resultList    
      });
    })
  }
  public render(): React.ReactElement<ISmartTableProps> {
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";  
    
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");  
   // SPComponentLoader.loadCss(cssURL);
    return (
      <div className={ styles.smartTable }>    
        <div className={ styles.container }>    
          <div className={ styles.row } style={{backgroundColor : this.props.color}}>    
            <div className={ styles.column }>    
              <span className={ styles.title }>Welcome to SharePoint!</span>    
              <p className={ styles.subTitle }>List items in react-bootstrap-table-next.</p>    
            </div>    
          </div>  
          {/* <div>
          <ListPicker context={this.props.context}  
              label="Select your list(s)"  
              placeHolder="Select your list(s)"  
              baseTemplate={100}  
              includeHidden={false}  
              multiSelect={false}  
              onSelectionChanged={this.onListPickerChange} /> 
          </div>  */}
          <br></br>
          {users.splice(0,users.length) && this.props.columnData && this.props.columnData.map(val => {
            users.push({
              dataField:val.internalName,
              text:val.displayName
            });
          })}
          <div className={styles.responsiveTable}>
          <BootstrapTable id="responsiveTable" keyField='id'
            data={this.state.projectMasterList}
            columns={ users.length === 0 ? projectMasterCol : users }
            bootstrap4 ={true}
            headerClasses="header-class"
            remote={ {
              filter: true,
              pagination: false,
              sort: false,
              cellEdit: false
            } }

            /> 
            </div>  
        </div>    
      </div>  
    );
  }
  private onListPickerChange (lists: string | string[]) {  
    console.log("Lists: ", lists);  
  }  
  private getAllProjectDetails = (): Promise<any> =>{    
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
   return web.lists.getById(this.props.lists).items.getAll().    
      then((results : any)=>{    
        return results;
      });       
  }
  private getFilteredProjectDetails = () =>{
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.lists.getById(this.props.lists).items.select('*')
    .filter(`substringof('',Title) or substringof('',ISDNNumber)`)
  }
  // private getFieldsForSelectedList =() =>{

  //   if(!this.props.lists){
  //     return Promise.resolve();
  //   }
  //   const filter = 'Hidden eq false and ReadOnlyField eq false';
  //   let web = new Web(this.props.context.pageContext.web.absoluteUrl);
  //   web.lists.getByTitle('Project Master').fields.select('Title').filter(filter)
  //   .get().then(data =>{
  //     console.log(data);
  //   });
  // }
 }
