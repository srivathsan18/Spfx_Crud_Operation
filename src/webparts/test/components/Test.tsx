import * as React from 'react';
import styles from './Test.module.scss';
import { ITestProps } from './ITestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, DatePicker, Button } from 'office-ui-fabric-react';

import { SPHttpClient,SPHttpClientResponse } from '@microsoft/sp-http';

// Test Purpose 


export default class Test extends React.Component<ITestProps, ITestState> {

  constructor(props: ITestProps) {
    super(props);
    this.state = {
     Title:"",
     ItemId:0,
     output:"Welcome"
    }
}

  public render(): React.ReactElement<ITestProps> { 
    return (  
    <div className={ styles.test }>
    <div className={ styles.container }>
      <div className={ styles.row }>
        <div className={ styles.column }>
      <div>
      <TextField
      label="Title"
      autoAdjustHeight
      value={this.state.Title}
      onChanged={val => {
        this.setState({ Title: val });
      }}
    />

{this.state.Title?
<a href="#" className={`${styles.button}`} onClick={() => this.createItem()}>  
                    <span className={styles.label}>Create item</span>  
                  </a>  
    
  :null}
  
  <div>
   
    <TextField
      label="ID"
      autoAdjustHeight
      value={this.state.ItemId.toString()}
      onChanged={val => {
        this.setState({ ItemId: val,output:"",Title:""});
      }}
    />
    <a href="#" className={`${styles.button}`} onClick={() => this.readItem()}>  
                    <span className={styles.label}>Read item</span>  
                  </a> 
   

    <a href="#" className={`${styles.button}`} onClick={() => this.updateItem()}>  
                    <span className={styles.label}>Update item</span>  
                  </a>   
                  <a href="#" className={`${styles.button}`} onClick={() => this.deleteItem()}>  
                    <span className={styles.label}>Delete item</span>  
                  </a> 
                  <div>
    {this.state.output}
    </div>
    </div>
    </div>
    </div>
    </div>
    </div>
    </div>
    );
  }

  private createItem(): void { 
    if(this.state.Title)
    {     
    const body: string = JSON.stringify({  
      'Title': this.state.Title
    });      
    this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: body  
    })  
    .then((response: SPHttpClientResponse) => {  
      return response.json();  
    })  
    .then((item): void => {  
      this.setState({  
        output:`Item ${item.Id} Created Successfully`,
      });
      alert(`ItemID ${item.Id} Created Successfully`);
         }, (error: any): void => {  
    }); 
  } 
  } 

  private readItem(): void {    
  
     this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${this.state.ItemId})?$select=Title`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          })

      .then((response: SPHttpClientResponse) => {  
        if(response.ok)
        {
        return response.json();  
        }
        else{
          return {Title:"Item Not Found"};
        }
      })  
      .then((item): void => {  
        this.setState({  
          output:item.Title
        });
      }, (error: any): void => {    
            
      });  
    
  }  

  private updateItem(): void {  
    
    if(this.state.Title)
    {
      if (!window.confirm('Are you sure you want to update the item?')) {  
        return;  
      }
  const body: string = JSON.stringify({  
          'Title': this.state.Title 
        });  

        this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${this.state.ItemId})`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'Content-type': 'application/json;odata=nometadata',  
              'odata-version': '',  
              'IF-MATCH': '*',  
              'X-HTTP-Method': 'MERGE'  
            },  
            body: body  
          })  
          .then((response: SPHttpClientResponse): void => {  
            if(response.ok)
            {
            this.setState({  
              output:`Item ${this.state.ItemId} Updated Successfully`
            });
          }
          else{
            this.setState({  
              output:'Item Not Found'
            });
          }
          }, (error: any): void => {  
           
          });  
  }  
}
private deleteItem(): void {  
  if (!window.confirm('Are you sure you want to delete the item?')) {  
    return;  
  }  
  
  let etag: string = undefined;  
 
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${this.state.ItemId})?$select=Id`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        })
    .then((response: SPHttpClientResponse) => {  
      etag = response.headers.get('ETag');  
      return response.json();  
    })  
    .then((item): Promise<SPHttpClientResponse> => {  
       
      return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=verbose',  
            'odata-version': '',  
            'IF-MATCH': etag,  
            'X-HTTP-Method': 'DELETE'  
          }  
        });  
    })  
    .then((response: SPHttpClientResponse): void => {  
      if(response.ok)
      {
      this.setState({  
        output: `Item ${this.state.ItemId} Deleted Successfully`,  
      });  
    }
    else{
      this.setState({  
        output:'Item Not Found'
      });
    }
    }, (error: any): void => {  
      
    });  
}  
}
