import * as React from 'react';
import styles from './CalendarWebpart.module.scss';
import { ICalendarWebpartProps } from './ICalendarWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp, List, ItemAddResult, ItemUpdateResult } from "@pnp/sp";
import { SPHttpClient, SPHttpClientResponse, MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';   



export default class CalendarWebpart extends React.Component<ICalendarWebpartProps, {}> {
  public getContact(){
    this.props.context.spHttpClient.get("https://lesh999.sharepoint.com/sites/feature-testing/_api/web/lists/getbytitle('TestList')/items(1)",SPHttpClient.configurations.v1,
    {
        headers:{
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': ''  
        }
    })
    .then((response:SPHttpClientResponse)=>{
        console.log(response.json());
    });
}
  public render(): React.ReactElement<ICalendarWebpartProps> {
    return (
      <div className={ styles.calendarWebpart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
              {/* <p>{this.props.context.pageContext.web.title}</p> */}
              <button onClick={(e)=>{
                debugger;
                this.props.context.msGraphClientFactory.getClient()
                .then((client:MSGraphClient)=>{
                  debugger
                  console.log(client);
                  client.api("/groups/88303b4a-2e14-40f9-8364-5472cae1c5b9/calendar/events")
                  .get((error, result:any, rawResponse?: any) => {
                    debugger;
                    console.log(result);
                });
                })
              }}>click</button>

                <button onClick={(e)=>{
                debugger;
                this.props.context.msGraphClientFactory.getClient()
                .then((client:MSGraphClient)=>{
                  debugger
                  console.log(client);
                  client.api("users")
                  .get()
                  .then((result:any)=>{
                    console.log(result);
                  })
                })
              }}>click2</button>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
