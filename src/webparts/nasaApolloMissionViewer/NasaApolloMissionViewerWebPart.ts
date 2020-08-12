import { 
  Version,
  Log 
} from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneTextFieldProps,
  PropertyPaneLabel,
  IPropertyPaneLabelProps,
  IPropertyPanePage
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NasaApolloMissionViewerWebPart.module.scss';
import * as strings from 'NasaApolloMissionViewerWebPartStrings';

import { IMission } from '../../models';
import { MissionService } from '../../services';
import { ButtonControl } from '../../services';

export interface INasaApolloMissionViewerWebPartProps {
  description: string;
  selectedMission: string;
}

export default class NasaApolloMissionViewerWebPart extends BaseClientSideWebPart <INasaApolloMissionViewerWebPartProps> {

  private selectedMission:IMission;
  private missionDetailElement: HTMLElement;

  protected onInit():Promise<void>{
    return new Promise<void>(
    (
      resolve: () => void,
      reject: (error:any) => void
    ): void => {
      this.selectedMission = this._getSelectedMission();
      resolve();
    }
  );
}

  public render(): void {
    Log.info('render()','This is INFO with service scope', this.context.serviceScope);
    Log.verbose('render()','This is VERBOSE with service scope', this.context.serviceScope);
    Log.warn('render()','This is WARN with service scope', this.context.serviceScope);
    //Log.error('render()',new Error('This is ERROR with service scope'), this.context.serviceScope);
    let webpartName :string = `[${this.context.webPartTag.replace(`.${this.context.instanceId}`,'')}]`;
    console.debug(webpartName,'Logging with web part prefix', this.selectedMission);
    console.table(this.selectedMission);
    console.table(this.selectedMission.crew);
    console.info(webpartName,'Logging with web part prefix');
    //console.warn('render().Console','This is for warn');
    //console.error('render().Console','This is for error');

    
    this.domElement.innerHTML = `
      <div class="${ styles.nasaApolloMissionViewer }">
       <div class="${ styles.container }">
         <div class="${ styles.row }">
          <div class="${ styles.column }">
            <span class="${ styles.title }">Welcome to SharePoint! This is updated version of NASA Apollo Mission solution (v2)</span>
            <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
                <a href="https://aka.ms/spfx" class="${ styles.button }">
                  <span class="${ styles.label }">Learn more</span>
                </a>
                <div class="apolloMissionDetails"></div>
              </div>
          </div>
       </div>
    </div>`;

    this.missionDetailElement = this.domElement.getElementsByClassName("apolloMissionDetails")[0] as HTMLElement;
    if(this.selectedMission){
      this._renderMissionDetails(this.missionDetailElement,this.selectedMission);
    } else{
      this.missionDetailElement.innerHTML = '';
    }
  }

  private _renderMissionDetails(element : HTMLElement, mission : IMission): void{
    element.innerHTML=`
      <p class="ms-font=m">
        <span class="ms-fontWeight-semibold">Mission: </span>
        ${escape(mission.name)}
      </p>
      <p class="ms-font=m">
        <span class="ms-fontWeight-semibold">Duration: </span>
        ${escape(this._getMissionTimeline(mission))}
      </p>
      <a href=${escape(mission.wiki_href)} class="${escape(styles.button)}" target="_blank">
        <span class="${styles.label}"> Learn more about ${escape(mission.name)} on wikipedia </span>
      </a>
    `;
  }
 
  protected get dataVersion(): Version {
  return Version.parse('1.0');
}

  private _getSelectedMission():IMission{
    const selectedMissionId:string = (this.properties.selectedMission)
    ? this.properties.selectedMission
    : "AS-506";
    return MissionService.getMission(selectedMissionId);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      //<IPropertyPanePage>ButtonControl.getPropertyPanePage(),
      {
        header:{
          description:'FIRST PAGE'
        },
        groups:[
          {
            groupName: 'First Group',
            groupFields: [
              PropertyPaneLabel('',{
                text: 'This is the label in first group on first page'
              })
            ]
          }
        ]
      },
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        displayGroupsAsAccordion:true,
        groups: [
          {
            groupName: strings.BasicGroupName,
            isCollapsed : true,
            groupFields: [
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel,
                disabled:true
              }),
              PropertyPaneTextField('selectedMission', {
                label: 'Mission Id',
                onGetErrorMessage: this._validateMissionCode.bind(this)
              })
            ]
          },
          {
            groupName: 'Group Second',
            isCollapsed : false,
            groupFields :[
              PropertyPaneLabel('',{
                text: 'This is label in second group',
                required: false
              })
            ]
          }
          //group 2
          //group 3
        ]
      }
    ]
  };
}
protected get disableReactivePropertyChanges(): boolean{
  return true;
}

protected onAfterPropertyPaneChangesApplied(): void{
  this.selectedMission = this._getSelectedMission();
  if(this.selectedMission){
  this._renderMissionDetails(this.missionDetailElement,this.selectedMission);
  }else{
    this.missionDetailElement.innerHTML ='';
  }
}

private _getMissionTimeline(mission:IMission):string{
  let missionDate = mission.end_date !== '' 
  ? `${escape(mission.launch_date.toString())} - ${escape(mission.end_date.toString())}`
  : `${escape(mission.launch_date.toString())}`; 
  return missionDate;
}

private _validateMissionCode(value:string):string{
  const validMissionCodeRegEx = /AS-[2,5][0-9][0-9]/g;
  return value.match(validMissionCodeRegEx) 
  ? ''
  : 'Invalid Mission Code - should be in format AS-###';
  
}


}
