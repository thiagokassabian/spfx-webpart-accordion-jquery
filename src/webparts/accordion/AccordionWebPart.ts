import { Version } from '@microsoft/sp-core-library';
import {
	IPropertyPaneConfiguration,
	IPropertyPaneDropdownOption,
	PropertyPaneDropdown,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AccordionWebPart.module.scss';
import * as strings from 'AccordionWebPartStrings';

import { IListInfo, sp } from '@pnp/sp/presets/all';
import * as $ from 'jquery';
import 'jqueryui';

require('../../../node_modules/jqueryui/jquery-ui.theme.min.css');

export interface IAccordionWebPartProps {
	title: string;
	selectionList: string;
}

export interface ISPList {
	Title: string;
	Texto: string;
}

export interface ISPLists {
	value: ISPList[];
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {
	private dropdownOptions: IPropertyPaneDropdownOption[];

	constructor() {
		super();
	}

	protected onInit(): Promise<void> {
		return super.onInit().then(_ => {
			sp.setup({
				spfxContext: this.context,
			});

			this._fetchOptions().then(response => (this.dropdownOptions = response));
		});
	}

	public render(): void {
		const accordionOptions: JQueryUI.AccordionOptions = {
			animate: true,
			heightStyle: 'content',
			collapsible: true,
			active: false,
			icons: false,
		};

		if (this.properties.selectionList) {
			this.domElement.innerHTML = `<div class="accordion ${styles.accordion}"/>`;
			this._getListData().then(response => {
				this._renderList(response.value);
				$('.accordion', this.domElement).accordion(accordionOptions);
			});
		} else {
			this.domElement.innerHTML = `
				<div>Selecione uma lista em configurações</div>
			`;
		}

		if (this.properties.title.trim()) {
			let title = document.createElement('h3');
			title.classList.add(styles.accordionTitle);
			title.innerHTML = `${this.properties.title.trim()}`;
			this.domElement.prepend(title);
		}
	}

	private _getListData(): Promise<ISPLists> {
		return sp.web.lists
			.getById(this.properties.selectionList)
			.items()
			.then(data => {
				let listData: ISPLists = { value: data };
				return listData;
			});
	}

	private _renderList(items: ISPList[]): void {
		let html: string = '';
		items.forEach((item: ISPList) => {
			console.log(`Título: ${item.Title} Conteúdo: ${item.Texto}`);
			html += `
			<h4 class="${styles.title}">${item.Title}</h4>
			<div class="${styles.content}">
				<div>${item.Texto}</div>
			</div>`;
		});

		const listContainer: Element = this.domElement.querySelector('.accordion');
		listContainer.innerHTML = html;
	}

	private _fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
		return sp.web.lists.get().then(response => {
			var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
			const filteredLists = response.filter(list => {
				return list.BaseTemplate == 100 && list.Hidden == false;
			});
			filteredLists.map((list: IListInfo) => {
				console.log('Found list with title = ' + list.Title);
				options.push({ key: list.Id, text: list.Title });
			});

			return options;
		});
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupFields: [
								PropertyPaneTextField('title', {
									label: 'Título (opcional)',
									value: '',
								}),
								PropertyPaneDropdown('selectionList', {
									label: 'Lista',
									options: this.dropdownOptions,
								}),
							],
						},
					],
				},
			],
		};
	}
}
