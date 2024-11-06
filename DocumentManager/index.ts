/* eslint-disable @typescript-eslint/no-explicit-any */
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import {DocumentManagerComponent} from './DocumentManagerComponent';
import { Buffer } from 'buffer';
import {getData} from './DynamicService';

export class DocumentManager implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    notifyOutputChanged: () => void;
    rootContainer: HTMLDivElement;
    selectedValue: string | null;
    context: ComponentFramework.Context<IInputs>;   
    private sasToken:string;
    private containerName:string;
    private _clientUrl:string;

    constructor()
    {
     //this.context.parameters.
    }


    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this.notifyOutputChanged = notifyOutputChanged;
        this.rootContainer = container;
        this.context = context; 
        window.Buffer = Buffer;
         console.log(context.parameters.ContainerName.raw);
        this._clientUrl = (
            window as any
          ).parent.Xrm.Utility.getGlobalContext().getClientUrl();
        getData(this._clientUrl, "environmentvariabledefinitions?$select=displayname,defaultvalue&$filter=displayname eq 'BlobSasToken' or displayname eq 'BlobContainer'").then((data)=>{
         this.sasToken = data.value.find((d:any) => d.displayname === 'BlobSasToken')?.defaultvalue;
         this.containerName = data.value.find((d: any) => d.displayname === 'BlobContainer')?.defaultvalue;
         const { value } = context.parameters;
         if (value && value.attributes) {
             ReactDOM.render(
                 React.createElement(DocumentManagerComponent, {
                     label: value.raw||'',
                     onChanges: this.onChange,
                     context:this.context,
                     sasToken:this.sasToken,
                     containerName:this.containerName
                 }),
                 this.rootContainer,
             );
         }
        })
       
    }

    onChange = (newValue: string | undefined): void => {
        this.context.parameters.value.raw = newValue||'';
        this.selectedValue = newValue||'';
       this.notifyOutputChanged();
  };
 
  

    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
      
    }


    public getOutputs(): IOutputs
    {
        return { value: this.selectedValue } as IOutputs;
    }

 
    public destroy(): void
    {
        ReactDOM.unmountComponentAtNode(this.rootContainer);
    }
}
