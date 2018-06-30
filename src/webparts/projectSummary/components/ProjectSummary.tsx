import * as React from 'react';
import styles from './ProjectSummary.module.scss';
import { IProjectSummaryProps } from './IProjectSummaryProps';
import { IProjectSummaryState } from './IProjectSummaryState';
import { escape } from '@microsoft/sp-lodash-subset';

import { ITableComponent } from './Table/Table';

//models
import { IProjectDataItem } from '../../../models'
import { IFieldItem } from '../../../models'
//services
import { ProjectService } from '../../../services'

export default class ProjectSummary extends React.Component<IProjectSummaryProps, IProjectSummaryState> {
  private projectService: ProjectService;

  constructor(props: IProjectSummaryProps) {
    super(props)

    this.state = { config: [], projectData: [] };

    this._getProjectData = this._getProjectData.bind(this)
  }

  componentDidMount() {
    this.projectService = new ProjectService(
      this.props.context.pageContext.web.absoluteUrl,
      this.props.context.spHttpClient
    )


    let _fieldNames: Array<IFieldItem> = this.trimFieldNames(this.props);
    this._getProjectData(_fieldNames)

  }

  componentWillReceiveProps(nextProps: IProjectSummaryProps, nextState: IProjectSummaryState) {
    let _fieldNames: Array<IFieldItem> = this.trimFieldNames(this.props);
    this._getProjectData(_fieldNames)

  }

  trimFieldNames(props: IProjectSummaryProps): Array<IFieldItem> {
    const _fieldPropertyNames = ['field1', 'field2', 'field3', 'field4', 'field5', 'field6', 'field7', 'field8', 'field9', 'field10']
    let _fieldNameArray: Array<IFieldItem> = [];
    _fieldPropertyNames.map((property, i) => {
      (typeof (props[property]) == 'string' && props[property].length > 0) ?
        // build up list of provided field names and add the display name
        _fieldNameArray.push({
          internalName: props[property],
          displayName: props[`column${i + 1}`] ? props[`column${i + 1}`] : '',
          fieldValue: ''
        }) :
        ''
    })
    return _fieldNameArray
  }

  private _getProjectData(fieldNames: Array<IFieldItem>): void {
    let _projectData: Array<IProjectDataItem> = []
    let _projectUrl = ''

    let _fieldNames: Array<IFieldItem> = fieldNames

    //used for lookup calls
    let _lookupApi = [];
    let _lookupData: Array<{ fieldName: string, fieldValueGuid: string }> = []

    this.projectService.getWebData()
      .then((result: any) => {
        console.log(result.id)
        _projectUrl = result.url
        return Promise.all([
          this.projectService.getCustomFields(result.url, result.id),
          this.projectService.getLookupTables(result.url, result.id),
          this.projectService.getProjectFields(result.url, result.id)
        ])
          .then((resultSet) => {
            console.log(resultSet)
            //lookup custom field names
            let _projectFields = this.projectService.processCustomFieldNames(resultSet[2], resultSet[0]['value'])
            //pull out required fields
            console.log(_projectFields)
            _fieldNames = this.projectService.getFieldValues(_fieldNames, _projectFields)
            console.log(_fieldNames)
            //if there are any lookup tables then lookup entry values

            _fieldNames.forEach((field, i) => {
              if (field.hasOwnProperty('fieldValue') && Array.isArray(field.fieldValue)) {
                //second entry in fieldvalue should be the custom field guid
                let _fieldId: string = field.fieldValue[1]
                //add api call to get associated lookup table entries for the lookuptable
                _lookupApi.push(this.projectService.getLookupEntryValues(_projectUrl, _fieldId))
                _lookupData.push({ fieldName: field.internalName, fieldValueGuid: field.fieldValue[0] })
              }
            })

            return Promise.all(_lookupApi)
          })
      })
      .then((results: Array<object>) => {
        if (results.length > 0) {
          results.forEach((entries, i) => {
            let _value = this.projectService.getLookupEntry(entries['value'], _lookupData[i])
            _fieldNames.forEach((field, u) => {
              if (field.internalName == _lookupData[i].fieldName) {
                _fieldNames[u].fieldValue = _value
              }
            })
          })
        }

        this.setState({
          config: [],
          projectData: _fieldNames
        })
      })
      .catch(err => console.error(err))

  }

  public render(): React.ReactElement<IProjectSummaryProps> {

    return (
      <div className={styles.projectSummary} >
        <ITableComponent config={this.state.config} projectData={this.state.projectData} />
      </div>
    );
  }
}
