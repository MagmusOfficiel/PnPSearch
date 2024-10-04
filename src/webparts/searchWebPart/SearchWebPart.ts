import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { Version, DisplayMode } from "@microsoft/sp-core-library";
import * as React from "react";
import * as ReactDom from "react-dom";
import {
    IDynamicDataAnnotatedPropertyValue,
    IDynamicDataCallables,
    IDynamicDataPropertyDefinition,
} from "@microsoft/sp-dynamic-data";
import { ISearchWebPartProps } from "./ISearchWebPartProps";
import PnPTelemetry from "@pnp/telemetry-js";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";
import * as strings from "SearchWebPartStrings";
import * as configs from "GlobalSetting";

export default class SearchWebPart
    extends BaseClientSideWebPart<ISearchWebPartProps>
    implements IDynamicDataCallables {
    private _availableConfigs: string[] = [];
    private _selectedConfig: string = "";
    private _isReloading: boolean = false;
    private sp: ReturnType<typeof spfi>;

    constructor() {
        super();
    }

    getPropertyDefinitions(): readonly IDynamicDataPropertyDefinition[] {
        throw new Error("Method not implemented.");
    }

    getPropertyValue(propertyId: string) {
        throw new Error("Method not implemented.");
    }

    getAnnotatedPropertyValue?(
        propertyId: string
    ): IDynamicDataAnnotatedPropertyValue {
        throw new Error("Method not implemented.");
    }

    protected async onInit(): Promise<void> {
        try {
            const telemetry = PnPTelemetry.getInstance();
            telemetry.optOut();
        } catch (error) {
            console.error("Error initializing:", error);
        }
        // Initialiser SPFx et charger les configurations disponibles
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
        return Version.parse("1.0");
    }

    public async render(): Promise<void> {
        if (this.displayMode === DisplayMode.Edit) {
            try {
                // Afficher l'interface de sélection de configuration
                this.renderConfigSelector();
            } catch (error) {
                console.error(`${strings.ErrorRendering}:`, error);
            }

            if (
                this.context.propertyPane &&
                this.context.propertyPane.isPropertyPaneOpen()
            ) {
                this.context.propertyPane.refresh();
            }
        } else {
            this.domElement.innerHTML = ""; // Masquer le webpart en mode lecture
        }
    }

    private async loadAvailableConfigs(): Promise<void> {
        try {
            const apiUrl = `${configs.Domain}${configs.BaseUrl}${configs.ApiWeb}('/${configs.BaseUrl}${configs.FolderConfig}')/Files`;
            const response = await this.fetchJson(apiUrl);
            this._availableConfigs = response.d.results.map(
                (file: any) => file.Name.split(".")[0]
            );
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
            const page = await this.sp.web.loadClientsidePage(
                this.context.pageContext.site.serverRequestPath
            );

            const searchWebParts = await this.sp.web.getClientsideWebParts();
            const webPartsConfig = [
                { config: searchBoxConfig, id: configs.SearchBoxId, defaultPosition: 1 , instanceId: configs.SearchBoxInstanceId },
                { config: searchFiltersConfig, id: configs.SearchFilterId, defaultPosition: 2, instanceId: configs.SearchFilterInstanceId },
                { config: searchVerticalConfig, id: configs.SearchVerticalId, defaultPosition: 3 ,instanceId: configs.SearchVerticalInstanceId },
                { config: searchResultConfig, id: configs.SearchResultId, defaultPosition: 4, instanceId: configs.SearchResultInstanceId }
            ];
    
            const sortedWebPartsConfig = webPartsConfig
                .map((wp, index) => ({
                    config: wp.config,
                    def: searchWebParts.find((part) => part.Id === wp.id),
                    position: wp.config?.position ?? wp.defaultPosition,
                    secondaryPosition: index,
                    instanceId: wp.instanceId
            
                }))
                .filter((wp) => wp.def && wp.config)
                .sort((a, b) => a.position - b.position || a.secondaryPosition - b.secondaryPosition);
    
            for (const { config, def, instanceId } of sortedWebPartsConfig) {
                const webPart = ClientsideWebpart.fromComponentDef(def);
                
                // Surcharger le `instanceId` généré automatiquement avec ton `instanceId` fixe
                webPart.data.webPartData.instanceId = instanceId;
                webPart.data.id = instanceId;
                const column = config.column ?? 1;
                if (config.verticalsDataSourceReference) {
                    config.verticalsDataSourceReference = `WebPart.${config.verticalsDataSourceReference.split('.')[1]}.${configs.SearchVerticalInstanceId}:pnpSearchVerticalsWebPart`;
                }
    
                if (config.filtersDataSourceReference) {
                    config.filtersDataSourceReference = `WebPart.${config.filtersDataSourceReference.split('.')[1]}.${configs.SearchFilterInstanceId}:pnpSearchFiltersWebPart`;
                }
    
                if (config.resultsDataSourceReference) {
                    config.resultsDataSourceReference = `WebPart.${config.resultsDataSourceReference.split('.')[1]}.${configs.SearchResultInstanceId}:pnpSearchResultsWebPart`;
                }
    
                if (config.boxDataSourceReference) {
                    config.boxDataSourceReference = `WebPart.${config.boxDataSourceReference.split('.')[1]}.${configs.SearchBoxInstanceId}:pnpSearchBoxWebPart`;
                }
                // Fonction pour obtenir l'instanceId en fonction de la propriété
                const getInstanceIdByProperty = (property: string): string => {
                    switch (property) {
                        case 'pnpSearchBoxWebPart':
                            return configs.SearchBoxInstanceId;
                        case 'pnpSearchFiltersWebPart':
                            return configs.SearchFilterInstanceId;
                        case 'pnpSearchVerticalsWebPart':
                            return configs.SearchVerticalInstanceId;
                        case 'pnpSearchResultsWebPart':
                            return configs.SearchResultInstanceId;
                        default:
                            return '';
                    }
                };
                // Mise à jour des références
                if (config.queryText && config.queryText.reference) {
                    const instanceIdProperty = getInstanceIdByProperty(config.queryText.reference._property);
                    config.queryText.reference._reference = `WebPart.${config.queryText.reference._reference.split('.')[1]}.${instanceIdProperty}:${config.queryText.reference._property}`;
                    config.queryText.reference._sourceId = `WebPart.${config.queryText.reference._sourceId.split('.')[1]}.${instanceIdProperty}`;
                    delete config.queryText.__type;
                }

                // Mise à jour des références
                if (config.selectedItemFieldValue && config.selectedItemFieldValue.reference) {
                    const instanceIdProperty = getInstanceIdByProperty(config.selectedItemFieldValue.reference._property);
                    config.selectedItemFieldValue.reference._reference = `WebPart.${config.selectedItemFieldValue.reference._reference.split('.')[1]}.${instanceIdProperty}:${config.selectedItemFieldValue.reference._property}`;
                    config.selectedItemFieldValue.reference._sourceId = `WebPart.${config.selectedItemFieldValue.reference._sourceId.split('.')[1]}.${instanceIdProperty}`;
                }

                webPart.setProperties(config);

                const existingSection = page.sections.find(section => section.columns.find(col => col.controls.find(control => control.id === webPart.data.webPartData.instanceId)));
                if (!existingSection) {
                    page.addSection().addColumn(column).addControl(webPart);
                }
            }

            await page.save();
            window.location.reload()
        } catch (error) {
            console.error(`${strings.ErrorRendering}:`, error);
        }
    }

    private async loadSelectedConfig(): Promise<void> {
        const baseUrl = `${configs.Domain}${configs.BaseUrl}${configs.FolderConfig}`;
        
        try {
            const apiUrl = `${baseUrl}/${this._selectedConfig}.json`;
            const selectedConfig = await this.fetchJson(apiUrl);

            if (
                selectedConfig &&
                selectedConfig["box"] &&
                selectedConfig["filter"] &&
                selectedConfig["vertical"] &&
                selectedConfig["result"]
            ) {
                await this.handleTemplate(
                    selectedConfig["box"],
                    selectedConfig["filter"],
                    selectedConfig["vertical"],
                    selectedConfig["result"]
                );
            } else {
                console.error(
                    "La configuration sélectionnée est incorrecte ou incomplète."
                );
            }
        } catch (error) {
            console.error(`${strings.ErrorLoadingConfigs}:`, error);
        }
    }

    private renderConfigSelector(): void {
        const element = React.createElement(
            "div",
            {
                style: {
                    textAlign: "center",
                    backgroundColor: "#f0f0f0",
                    padding: "20px",
                    borderRadius: "5px",
                },
            },
            React.createElement("h2", {}, strings.SelectConfiguration),
            React.createElement(
                "select",
                {
                    onChange: this.selectedConfigName.bind(this),
                    style: { margin: "10px", padding: "5px" },
                },
                React.createElement(
                    "option",
                    { key: "default-config", value: "" },
                    strings.SelectConfiguration
                ),
                this._availableConfigs.map((config) =>
                    React.createElement(
                        "option",
                        { key: config, value: config },
                        `${config}`
                    )
                )
            ),
            React.createElement(
                "button",
                {
                    onClick: this.reloadConfiguration.bind(this),
                    style: {
                        margin: "10px",
                        padding: "10px 20px",
                        backgroundColor: "#d0d0d0",
                        border: "none",
                        borderRadius: "5px",
                        cursor: this._isReloading ? "not-allowed" : "pointer",
                        position: "relative",
                    },
                    disabled: this._isReloading,
                },
                strings.ReloadConfiguration
            )
        );

        ReactDom.render(element, this.domElement);
    }

    private async reloadConfiguration(): Promise<void> {
        if (this._isReloading) {
            // Empêcher les appels multiples en parallèle
            return;
        }
    
        this._isReloading = true; // Verrouiller le processus de rechargement
        try {
            // Mettre à jour l'interface utilisateur pour refléter le rechargement en cours
            this.renderConfigSelector();
            
            // Vérification si une configuration est sélectionnée avant de charger
            if (!this._selectedConfig) {
                return;
            }
    
            // Charger la configuration sélectionnée
            await this.loadSelectedConfig();
        } catch (error) {
            console.error(`${strings.ErrorLoadingConfigs}:`, error);
        } finally {
            // Réinitialiser l'état de rechargement et mettre à jour l'interface utilisateur
            this._isReloading = false;
            this.renderConfigSelector();
        }
    }
    

    private selectedConfigName(
        event: React.ChangeEvent<HTMLSelectElement>
    ): void {
        this._selectedConfig = event.target.value;
    }

    private async fetchJson(url: string): Promise<any> {
        const response = await fetch(url, {
            headers: {
                Accept: "application/json;odata=verbose",
            },
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        return response.json();
    }
}
