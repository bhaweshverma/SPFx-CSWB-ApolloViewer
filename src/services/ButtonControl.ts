import {
    PropertyPaneButton,
    PropertyPaneButtonType,
    IPropertyPaneButtonProps,
    IPropertyPanePage,
    PropertyPaneHorizontalRule,
    PropertyPaneLabel,
    IPropertyPaneLabelProps
 } from '@microsoft/sp-property-pane';

 export class ButtonControl{
     public onButtonClick(){
         alert('Button clicked');
     }
     public getPropertyPanePage(): IPropertyPanePage{
         return <IPropertyPanePage>{
            header:{
                description:''
            },
            groups:[
                {
                    groupName:'Button Group',
                    groupFields: [
                        PropertyPaneLabel('',{
                            text: 'Label for Button'
                        }),
                        PropertyPaneButton('',{
                            text: 'Apollo Button',
                            onClick: this.onButtonClick
                        })
                    ]
                }
            ]
         };
     }
 }