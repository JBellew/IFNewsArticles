"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * Table settings
 */
class TableSettingsCard extends FormattingSettingsCard {

    headerColor = new formattingSettings.ColorPicker({
        name: "headerColor",
        displayName: "Header Color",
        value: { value: "#000000" }
    });

    headerFontColor = new formattingSettings.ColorPicker({
        name: "headerFontColor",
        displayName: "Header Font Color",
        value: { value: "#FFFFFF" }
    });

    tableHeight = new formattingSettings.NumUpDown({
        name: "tableHeight",
        displayName: "Table Height (vh)",
        value: 50
    });

    pageLength = new formattingSettings.NumUpDown({
        name: "pageLength",
        displayName: "Page Length",
        value: 10
    });
    

    name: string = "tableSettings";
    displayName: string = "Table Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.headerColor,
        this.headerFontColor,
        this.tableHeight,
        this.pageLength
    ];
}

/**
 * Visual settings model class
 */
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    tableSettingsCard = new TableSettingsCard();
    cards = [this.tableSettingsCard];
}
