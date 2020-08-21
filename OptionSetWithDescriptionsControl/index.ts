import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IDropdownOptionExt, IOptionSetWithDescriptionsProperties, OptionSetControl } from './OptionSetControl';

export class OptionSetWithDescriptionsControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private currentValue: number | null;
	private container: HTMLDivElement;
	private notifyOutputChanged: () => void;
	private isDisabled: boolean;
	private defaultValue : number | undefined;
	private context: ComponentFramework.Context<IInputs>;	
	private optionSetProperties: IOptionSetWithDescriptionsProperties;

	constructor()
	{

	}	

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this.container = container;
		this.notifyOutputChanged = notifyOutputChanged;
		this.renderControl(context);
		this.context = context;
		this.defaultValue = context.parameters.optionSet.attributes?.DefaultValue;
		
	}

	private async getMetadata(): Promise<IDropdownOptionExt[]> {
		const version = '9.0';
		const webapiurl = '/api/data/v' + version + '/';
		//@ts-ignore
		const entityName = this.context.mode.contextInfo.entityTypeName;
		const boundFieldName = this.context.parameters.optionSet.attributes?.LogicalName ?? '';
        const optionSetQueryUrl = webapiurl + "EntityDefinitions(LogicalName='" + entityName + 
        "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata" + 
		"?$select=LogicalName&$filter=LogicalName eq '" + boundFieldName + "'&$expand=OptionSet";
		
		let response = await fetch(optionSetQueryUrl);
		let data = await response.json();
		let items: IDropdownOptionExt[] = [];
		data.value[0].OptionSet.Options.forEach(function (option: any, i: number) {
			var property: IDropdownOptionExt = {
				"text": option.Label.UserLocalizedLabel.Label,
				"color": option.Color,
				"description": option.Description.UserLocalizedLabel.Label,
				"key": option.Value
			}
			items.push(property);
		});

		return items;
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.renderControl(context);
	}
	
	private renderControl(context: ComponentFramework.Context<IInputs>): void {		
		this.getMetadata()
		.then(result => {
			this.isDisabled = context.mode.isControlDisabled;
			this.currentValue = context.parameters.optionSet.raw;		
			this.optionSetProperties = {
				//@ts-ignore
				entityName: context.mode.contextInfo.entityTypeName,
				boundFieldName: context.parameters.optionSet.attributes?.LogicalName ?? '',
				selectedKey: this.currentValue, 
				onChange: (newValue: number |null) => {
					this.currentValue = newValue;
					this.notifyOutputChanged();
				},
				isDisabled : this.isDisabled, 
				defaultValue : this.defaultValue,
				optionsetItems : result
			}
			ReactDOM.render(React.createElement(OptionSetControl, this.optionSetProperties), this.container);
		});		
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			optionSet: this.currentValue == null ? undefined : this.currentValue
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		ReactDOM.unmountComponentAtNode(this.container);
	}
}