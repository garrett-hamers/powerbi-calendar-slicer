/*
 *  Power BI Visual — Atlyn Calendar Slicer
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";

export class Visual implements IVisual {
    private readonly target: HTMLElement;
    private readonly root: HTMLElement;
    private readonly host: IVisualHost;
    private readonly localizationManager: ILocalizationManager;
    private readonly formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;

    private readonly isHighContrast: boolean;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.localizationManager = this.host.createLocalizationManager();
        this.formattingSettingsService =
            new FormattingSettingsService(this.localizationManager);
        this.formattingSettings = new VisualFormattingSettingsModel();

        this.isHighContrast = this.host.colorPalette.isHighContrast === true;

        this.root = document.createElement("div");
        this.root.className = "atlynCalendarSlicer";
        this.root.classList.toggle("high-contrast", this.isHighContrast);
        this.root.setAttribute("role", "group");
        this.root.setAttribute(
            "aria-label",
            this.localize("Aria_Calendar", "Calendar date slicer")
        );
        this.target.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        this.host.eventService?.renderingStarted(options);
        try {
            this.clear();

            const dataView: DataView | undefined = options.dataViews?.[0];
            this.formattingSettings =
                this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    dataView
                );

            if (!dataView?.categorical?.categories?.length) {
                this.renderLanding(
                    this.localize("Landing_AddField", "Add a date field to the Date bucket")
                );
            }

            this.host.eventService?.renderingFinished(options);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.host.eventService?.renderingFailed(options, message);
            throw error;
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private clear(): void {
        while (this.root.firstChild) {
            this.root.removeChild(this.root.firstChild);
        }
    }

    private renderLanding(message: string): void {
        const landing = document.createElement("div");
        landing.className = "cs-landing";

        const title = document.createElement("div");
        title.className = "cs-landing-title";
        title.textContent = this.localize("Visual_Title", "Atlyn Calendar Slicer");
        landing.appendChild(title);

        const body = document.createElement("div");
        body.textContent = message;
        landing.appendChild(body);

        this.root.appendChild(landing);
    }

    /**
     * Resolve a localized string, falling back to English when the host echoes
     * the key back (as some test hosts do) or the key is unknown.
     */
    private localize(key: string, fallback: string): string {
        const resolved = this.localizationManager?.getDisplayName(key);
        return resolved && resolved !== key ? resolved : fallback;
    }
}
