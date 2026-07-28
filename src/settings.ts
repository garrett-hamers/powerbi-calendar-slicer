/**
 * Formatting settings for the Atlyn Calendar Slicer.
 *
 * Every card name here matches an `objects` key in capabilities.json and every
 * slice name matches a property name, so the format pane and the persisted
 * object model stay in lock-step.
 */
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class CalendarCard extends FormattingSettingsCard {
    weekStartDay = new formattingSettings.ItemDropdown({
        name: "weekStartDay",
        displayName: "Week Starts On",
        displayNameKey: "Calendar_WeekStartDay",
        items: [
            { value: "0", displayName: "Sunday" },
            { value: "1", displayName: "Monday" },
            { value: "6", displayName: "Saturday" }
        ],
        value: { value: "0", displayName: "Sunday" }
    });

    monthsToShow = new formattingSettings.NumUpDown({
        name: "monthsToShow",
        displayName: "Months Shown",
        displayNameKey: "Calendar_MonthsToShow",
        description: "Number of months to display side by side",
        descriptionKey: "Calendar_MonthsToShow_Desc",
        value: 1,
        options: {
            minValue: { type: 0, value: 1 },
            maxValue: { type: 1, value: 4 }
        }
    });

    showTodayMarker = new formattingSettings.ToggleSwitch({
        name: "showTodayMarker",
        displayName: "Mark Today",
        displayNameKey: "Calendar_ShowTodayMarker",
        value: true
    });

    showWeekNumbers = new formattingSettings.ToggleSwitch({
        name: "showWeekNumbers",
        displayName: "Show Week Numbers",
        displayNameKey: "Calendar_ShowWeekNumbers",
        description: "Show ISO-8601 week numbers in a leading column",
        descriptionKey: "Calendar_ShowWeekNumbers_Desc",
        value: false
    });

    fiscalYearStartMonth = new formattingSettings.ItemDropdown({
        name: "fiscalYearStartMonth",
        displayName: "Fiscal Year Starts",
        displayNameKey: "Calendar_FiscalYearStartMonth",
        description: "First month of the fiscal year; drives the QTD and YTD presets",
        descriptionKey: "Calendar_FiscalYearStartMonth_Desc",
        items: [
            { value: "1", displayName: "January" },
            { value: "2", displayName: "February" },
            { value: "3", displayName: "March" },
            { value: "4", displayName: "April" },
            { value: "5", displayName: "May" },
            { value: "6", displayName: "June" },
            { value: "7", displayName: "July" },
            { value: "8", displayName: "August" },
            { value: "9", displayName: "September" },
            { value: "10", displayName: "October" },
            { value: "11", displayName: "November" },
            { value: "12", displayName: "December" }
        ],
        value: { value: "1", displayName: "January" }
    });

    name = "calendar";
    displayName = "Calendar";
    displayNameKey = "Card_Calendar";
    slices: FormattingSettingsSlice[] = [
        this.weekStartDay,
        this.monthsToShow,
        this.showTodayMarker,
        this.showWeekNumbers,
        this.fiscalYearStartMonth
    ];
}

class CellsCard extends FormattingSettingsCard {
    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Text Colour",
        displayNameKey: "Cells_TextColor",
        value: { value: "#333333" }
    });

    headerColor = new formattingSettings.ColorPicker({
        name: "headerColor",
        displayName: "Header Colour",
        displayNameKey: "Cells_HeaderColor",
        value: { value: "#666666" }
    });

    selectedColor = new formattingSettings.ColorPicker({
        name: "selectedColor",
        displayName: "Selected Colour",
        displayNameKey: "Cells_SelectedColor",
        value: { value: "#0078D4" }
    });

    todayColor = new formattingSettings.ColorPicker({
        name: "todayColor",
        displayName: "Today Outline",
        displayNameKey: "Cells_TodayColor",
        value: { value: "#0078D4" }
    });

    weekendShading = new formattingSettings.ToggleSwitch({
        name: "weekendShading",
        displayName: "Shade Weekends",
        displayNameKey: "Cells_WeekendShading",
        value: false
    });

    weekendColor = new formattingSettings.ColorPicker({
        name: "weekendColor",
        displayName: "Weekend Colour",
        displayNameKey: "Cells_WeekendColor",
        value: { value: "#F6F6F6" }
    });

    name = "cells";
    displayName = "Cells";
    displayNameKey = "Card_Cells";
    slices: FormattingSettingsSlice[] = [
        this.textColor,
        this.headerColor,
        this.selectedColor,
        this.todayColor,
        this.weekendShading,
        this.weekendColor
    ];
}

class HeatmapCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Heat-shade Days",
        displayNameKey: "Heatmap_Show",
        description: "Colour each day by the Values measure. Binding a measure may disable slicer sync across pages.",
        descriptionKey: "Heatmap_Show_Desc",
        value: false
    });

    minColor = new formattingSettings.ColorPicker({
        name: "minColor",
        displayName: "Low Colour",
        displayNameKey: "Heatmap_MinColor",
        value: { value: "#DEEBF7" }
    });

    maxColor = new formattingSettings.ColorPicker({
        name: "maxColor",
        displayName: "High Colour",
        displayNameKey: "Heatmap_MaxColor",
        value: { value: "#08519C" }
    });

    datesWithDataOnly = new formattingSettings.ToggleSwitch({
        name: "datesWithDataOnly",
        displayName: "Grey Days Without Data",
        displayNameKey: "Heatmap_DatesWithDataOnly",
        value: false
    });

    name = "heatmap";
    displayName = "Heatmap";
    displayNameKey = "Card_Heatmap";
    slices: FormattingSettingsSlice[] = [
        this.show,
        this.minColor,
        this.maxColor,
        this.datesWithDataOnly
    ];
}

class PresetsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show Preset Buttons",
        displayNameKey: "Presets_Show",
        value: true
    });

    name = "presets";
    displayName = "Presets";
    displayNameKey = "Card_Presets";
    slices: FormattingSettingsSlice[] = [this.show];
}

class InteractionCard extends FormattingSettingsCard {
    multiSelect = new formattingSettings.ToggleSwitch({
        name: "multiSelect",
        displayName: "Allow Multi-select",
        displayNameKey: "Interaction_MultiSelect",
        value: true
    });

    name = "interaction";
    displayName = "Interaction";
    displayNameKey = "Card_Interaction";
    slices: FormattingSettingsSlice[] = [this.multiSelect];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    calendarCard = new CalendarCard();
    cellsCard = new CellsCard();
    heatmapCard = new HeatmapCard();
    presetsCard = new PresetsCard();
    interactionCard = new InteractionCard();

    cards: FormattingSettingsCard[] = [
        this.calendarCard,
        this.cellsCard,
        this.heatmapCard,
        this.presetsCard,
        this.interactionCard
    ];
}
