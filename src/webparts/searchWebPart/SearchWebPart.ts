import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IDynamicDataAnnotatedPropertyValue, IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { ISearchWebPartProps } from './ISearchWebPartProps';
import PnPTelemetry from '@pnp/telemetry-js';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";
import * as strings from 'SearchWebPartStrings';
import * as configs from 'GlobalSetting';

export default class SearchWebPart extends BaseClientSideWebPart<ISearchWebPartProps> implements IDynamicDataCallables {
    private _availableConfigs: string[] = [];
    private _selectedConfig: string = '';
    private _isReloading: boolean = false;
    private sp: ReturnType<typeof spfi>;

    constructor() {
        super();
    }

    getPropertyDefinitions(): readonly IDynamicDataPropertyDefinition[] {
        throw new Error('Method not implemented.');
    }

    getPropertyValue(propertyId: string) {
        throw new Error('Method not implemented.');
    }

    getAnnotatedPropertyValue?(propertyId: string): IDynamicDataAnnotatedPropertyValue {
        throw new Error('Method not implemented.');
    }

    protected async onInit(): Promise<void> {
        try {
            const telemetry = PnPTelemetry.getInstance();
            telemetry.optOut();
        } catch (error) {
            console.error('Error initializing:', error);
        }

        await this.loadAvailableConfigs();

        this.sp = spfi().using(SPFx(this.context));

        return super.onInit();
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get isRenderAsync(): boolean {
        return true;
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    public async render(): Promise<void> {
        if (this.displayMode === DisplayMode.Edit) {
            try {
                await this.loadAvailableConfigs();
                this.renderConfigSelector();
            } catch (error) {
                console.error(`${strings.ErrorRendering}:`, error);
            }

            if (this.context.propertyPane && this.context.propertyPane.isPropertyPaneOpen()) {
                this.context.propertyPane.refresh();
            }
        } else {
            this.domElement.innerHTML = '';  // Masquer le webpart en mode lecture
        }
    }

    private async loadAvailableConfigs(): Promise<void> {
        try {
            const response = await this.fetchJson(`${configs.Domain + configs.BaseUrl + configs.ApiWeb}('/${configs.BaseUrl + configs.FolderConfig}')/Files`);
            this._availableConfigs = response.d.results.map((file: any) => file.Name.split('.')[0]);
        } catch (error) {
            console.error(`${strings.ErrorLoadingConfigs}:`, error);
        }
    }

    private async handleTemplate(
        searchBoxConfig?: any,
        searchFiltersConfig?: any,
        searchVerticalConfig?: any,
        searchResultConfig?: any
    ): Promise<void> {
        try {
            const page = await this.sp.web.loadClientsidePage(this.context.pageContext.site.serverRequestPath);
            const searchWebParts = await this.sp.web.getClientsideWebParts();
            const webPartsConfig = [
                { config: searchBoxConfig, id: configs.SearchBoxId, defaultPosition: 1 },
                { config: searchFiltersConfig, id: configs.SearchFilterId, defaultPosition: 2 },
                { config: searchVerticalConfig, id: configs.SearchVerticalId, defaultPosition: 3 },
                { config: searchResultConfig, id: configs.SearchResultId, defaultPosition: 4 },
            ];

            const sortedWebPartsConfig = webPartsConfig
                .map((wp, index) => ({
                    config: wp.config,
                    def: searchWebParts.find(part => part.Id === wp.id),
                    position: wp.config ? (wp.config['position'] ?? wp.defaultPosition) : wp.defaultPosition,
                    secondaryPosition: index
                }))
                .filter(wp => wp.def && wp.config) // Filtrer les configurations et définitions définies
                .sort((a, b) => a.position - b.position || a.secondaryPosition - b.secondaryPosition);

            for (const { config, def } of sortedWebPartsConfig) {
                const webPart = ClientsideWebpart.fromComponentDef(def);
                const column = config['column'] ?? 1;
                webPart.setProperties(config);
                page.addSection().addColumn(column).addControl(webPart);
            }

            await page.save();

            this._isReloading = false;
            window.location.reload();

        } catch (error) {
            console.error(`${strings.ErrorRendering}:`, error);
        }
    }

    private async loadSelectedConfig(): Promise<void> {
        const baseUrl = configs.Domain + configs.BaseUrl + configs.FolderConfig;
        try {
            const selectedConfig = await this.fetchJson(`${baseUrl}/${this._selectedConfig}.json`);
            await this.handleTemplate(selectedConfig['box'], selectedConfig['filter'], selectedConfig['vertical'], selectedConfig['result']);
        } catch (error) {
            console.error(`${strings.ErrorLoadingConfigs}:`, error);
        }
    }

    private renderConfigSelector(): void {
        const element = React.createElement('div', { style: { textAlign: 'center', backgroundColor: '#f0f0f0', padding: '20px', borderRadius: '5px' } },
            React.createElement('h2', {}, strings.SelectConfiguration),
            React.createElement('select', { onChange: this.selectedConfigName.bind(this), style: { margin: '10px', padding: '5px' } },
                React.createElement('option', { key: 'default-config', value: '' }, strings.SelectConfiguration),
                this._availableConfigs.map((config) =>
                    React.createElement('option', { key: config, value: config }, `${config}`)
                )
            ),
            React.createElement('button', {
                onClick: this.reloadConfiguration.bind(this),
                style: {
                    margin: '10px', padding: '10px 20px', backgroundColor: '#d0d0d0',
                    border: 'none', borderRadius: '5px', cursor: this._isReloading ? 'not-allowed' : 'pointer',
                    position: 'relative'
                },
                disabled: this._isReloading
            },
                strings.ReloadConfiguration
            )
        );

        ReactDom.render(element, this.domElement);
    }

    private async reloadConfiguration(): Promise<void> {
        if (this._isReloading) {
            return;
        }
        this._isReloading = true;
        this.renderConfigSelector();  
        await this.loadSelectedConfig();
        this._isReloading = false;
        this.renderConfigSelector();  
    }

    private selectedConfigName(event: React.ChangeEvent<HTMLSelectElement>): void {
        this._selectedConfig = event.target.value;
    }

    private async fetchJson(url: string): Promise<any> {
        const response = await fetch(url, {
            headers: {
                'Accept': 'application/json;odata=verbose'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        return response.json();
    }
}
