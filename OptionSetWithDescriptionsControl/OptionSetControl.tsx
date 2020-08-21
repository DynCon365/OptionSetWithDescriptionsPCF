import * as React from 'react';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { Stack } from '@fluentui/react/lib/Stack';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import {
    TooltipHost,
    TooltipDelay,
    DirectionalHint,
    ITooltipHostStyles,
  } from 'office-ui-fabric-react/lib/Tooltip';
import { useId } from '@uifabric/react-hooks';

initializeIcons();

export interface IOptionSetWithDescriptionsProperties {
    entityName: string;    
    boundFieldName: string;           
    onChange: (value: number|null) => void;
    selectedKey: number | null;    
    isDisabled : boolean;
    defaultValue : number | undefined;
    optionsetItems: IDropdownOptionExt[];
}

export interface IDropdownOptionExt extends IDropdownOption {
    description: string;
    color: string;
}

const calloutProps = { gapSpace: 5 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline' } }; 

export const OptionSetControl: React.FC<IOptionSetWithDescriptionsProperties> = (props: IOptionSetWithDescriptionsProperties) => {
    const tooltipId = useId('tooltip');   
    const onDropdownChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        if (item) {       
            //setToolTipText(props.optionsetItems.find(e => e.key === item.key)?.description ?? '');   
            props.onChange(item.key as number);
        }       
    };
   
    return (
        <TooltipHost
        content={props.optionsetItems.find(e => e.key === props.selectedKey)?.description ?? ''}
        id={tooltipId}
        calloutProps={calloutProps}
        styles={hostStyles}
        delay={TooltipDelay.zero}
        directionalHint={DirectionalHint.bottomAutoEdge}
        >
        <Stack>            
            <Dropdown
                placeHolder="---"
                aria-describedby={tooltipId}
                selectedKey={props.selectedKey}
                defaultValue={props.defaultValue}
                disabled={props.isDisabled}
                options={props.optionsetItems}
                onChange={onDropdownChange}
            />           
        </Stack>        
      </TooltipHost>        
    )
};