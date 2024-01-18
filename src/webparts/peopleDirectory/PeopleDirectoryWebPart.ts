import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneCheckbox, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { PeopleDirectory, IPeopleDirectoryProps } from './components/PeopleDirectory/';
import * as strings from 'PeopleDirectoryWebPartStrings';

export interface IPeopleDirectoryWebPartProps {
	title: string;
	additionalFilterKQL: string;
	indexByLastname:boolean;
}

export default class PeopleDirectoryWebPart extends BaseClientSideWebPart<IPeopleDirectoryWebPartProps> {

	protected onInit(): Promise<void> {
		document.documentElement.style
			.setProperty('--maxPersonaWidth', this.width > 640 ? "50%" : "100");
		return Promise.resolve();
	}
	protected onAfterResize(newWidth: number) {
		
		document.documentElement.style
			.setProperty('--maxPersonaWidth', newWidth > 640 ? "50%" : "100");
	}
	public render(): void {
		const element: React.ReactElement<IPeopleDirectoryProps> = React.createElement(
			PeopleDirectory,
			{
				webUrl: this.context.pageContext.web.absoluteUrl,
				spHttpClient: this.context.spHttpClient,
				title: this.properties.title,
				displayMode: this.displayMode,
				locale: this.getLocaleId(),
				additionalFilterKQL: this.properties.additionalFilterKQL,
				indexByLastname:this.properties.indexByLastname,
				onTitleUpdate: (newTitle: string) => {
					// after updating the web part title in the component
					// persist it in web part properties
					this.properties.title = newTitle;
				}
			}
		);

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getLocaleId(): string {
		return this.context.pageContext.cultureInfo.currentUICultureName;
	}
	protected get disableReactivePropertyChanges(): boolean {
		return true;
	}
	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [{
				header: {
					description: strings.PropertyPaneDescription
				},
				groups: [
					{
						groupName: strings.SettingsLabel,
						groupFields: [
							PropertyPaneCheckbox('indexByLastname', {
								text:strings.IndexByLastname,
								checked:true
							})
						]
					},
					{
						groupName: strings.FilterSettingsLabel,
						groupFields: [
							PropertyPaneTextField('additionalFilterKQL', {
								label: strings.AdditionalFilterFieldLabel,
								description: strings.AdditionalFilterFieldDescription
							})
						]
					}
				]
			}]
		};
	}
}
