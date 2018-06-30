import * as React from 'react';
import styles from "../ProjectSummary.module.scss";

import { ITableProps } from './ITableProps';
import { ITableState } from './ITableState';
import { IFieldItem } from '../../../../models'


export class ITableComponent extends React.Component<ITableProps, ITableState> {
    constructor(props: ITableProps) {
        super(props);

        this.state = { projectData: [] };

    }

    componentWillReceiveProps(newProps: ITableProps, newState: ITableState) {
        console.log(newProps);
        this.setState({
            projectData: newProps.projectData
        })
    }

    public render(): React.ReactElement<ITableProps> {

        return (
            <div>
                {
                    this.state.projectData.map((item: IFieldItem, i) => {
                        return (<div key={i}>

                            <span><strong>{item.displayName}</strong>  :  {item.fieldValue}</span>

                        </div>
                        )
                    })
                }
            </div>
        )
    }


}