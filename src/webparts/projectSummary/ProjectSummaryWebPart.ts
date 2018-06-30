import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneTextFieldProps
} from '@microsoft/sp-webpart-base';

import * as strings from 'ProjectSummaryWebPartStrings';
import ProjectSummary from './components/ProjectSummary';
import { IProjectSummaryProps } from './components/IProjectSummaryProps';



export interface IProjectSummaryWebPartProps {
  field1: string;
  field2: string;
  field3: string;
  field4: string;
  field5: string;
  field6: string;
  field7: string;
  field8: string;
  field9: string;
  field10: string;
  column1: string;
  column2: string;
  column3: string;
  column4: string;
  column5: string;
  column6: string;
  column7: string;
  column8: string;
  column9: string;
  column10: string;
}

export default class ProjectSummaryWebPart extends BaseClientSideWebPart<IProjectSummaryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
      ProjectSummary,
      {
        field1: this.properties.field1,
        field2: this.properties.field2,
        field3: this.properties.field3,
        field4: this.properties.field4,
        field5: this.properties.field5,
        field6: this.properties.field6,
        field7: this.properties.field7,
        field8: this.properties.field8,
        field9: this.properties.field9,
        field10: this.properties.field10,
        column1: this.properties.column1,
        column2: this.properties.column2,
        column3: this.properties.column3,
        column4: this.properties.column4,
        column5: this.properties.column5,
        column6: this.properties.column6,
        column7: this.properties.column7,
        column8: this.properties.column8,
        column9: this.properties.column9,
        column10: this.properties.column10,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied() {

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.Pane1Description
          },
          groups: [
            {
              groupName: strings.GroupFieldName,
              groupFields: [
                PropertyPaneTextField('field1', <IPropertyPaneTextFieldProps>{
                  label: strings.Field1Label
                }),
                PropertyPaneTextField('field2', <IPropertyPaneTextFieldProps>{
                  label: strings.Field2Label
                }),
                PropertyPaneTextField('field3', <IPropertyPaneTextFieldProps>{
                  label: strings.Field3Label
                }),
                PropertyPaneTextField('field4', <IPropertyPaneTextFieldProps>{
                  label: strings.Field4Label
                }),
                PropertyPaneTextField('field5', <IPropertyPaneTextFieldProps>{
                  label: strings.Field5Label
                }),
                PropertyPaneTextField('field6', <IPropertyPaneTextFieldProps>{
                  label: strings.Field6Label
                }),
                PropertyPaneTextField('field7', <IPropertyPaneTextFieldProps>{
                  label: strings.Field7Label
                }),
                PropertyPaneTextField('field8', <IPropertyPaneTextFieldProps>{
                  label: strings.Field8Label
                }),
                PropertyPaneTextField('field9', <IPropertyPaneTextFieldProps>{
                  label: strings.Field9Label
                }),
                PropertyPaneTextField('field10', <IPropertyPaneTextFieldProps>{
                  label: strings.Field10Label
                })

              ]
            },
            {
              groupName: strings.GroupColumnHeading,
              groupFields: [
                PropertyPaneTextField('column1', <IPropertyPaneTextFieldProps>{
                  label: strings.Column1HeadingLabel
                }),
                PropertyPaneTextField('column2', <IPropertyPaneTextFieldProps>{
                  label: strings.Column2HeadingLabel
                }),
                PropertyPaneTextField('column3', <IPropertyPaneTextFieldProps>{
                  label: strings.Column3HeadingLabel
                }),
                PropertyPaneTextField('column4', <IPropertyPaneTextFieldProps>{
                  label: strings.Column4HeadingLabel
                }),
                PropertyPaneTextField('column5', <IPropertyPaneTextFieldProps>{
                  label: strings.Column5HeadingLabel
                }),
                PropertyPaneTextField('column6', <IPropertyPaneTextFieldProps>{
                  label: strings.Column6HeadingLabel
                }),
                PropertyPaneTextField('column7', <IPropertyPaneTextFieldProps>{
                  label: strings.Column7HeadingLabel
                }),
                PropertyPaneTextField('column8', <IPropertyPaneTextFieldProps>{
                  label: strings.Column8HeadingLabel
                }),
                PropertyPaneTextField('column9', <IPropertyPaneTextFieldProps>{
                  label: strings.Column9HeadingLabel
                }),
                PropertyPaneTextField('column10', <IPropertyPaneTextFieldProps>{
                  label: strings.Column10HeadingLabel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
