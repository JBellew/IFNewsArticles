import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;
/**
 * Table settings
 */
declare class TableSettingsCard extends FormattingSettingsCard {
    headerColor: formattingSettings.ColorPicker;
    headerFontColor: formattingSettings.ColorPicker;
    tableHeight: formattingSettings.NumUpDown;
    pageLength: formattingSettings.NumUpDown;
    name: string;
    displayName: string;
    slices: Array<FormattingSettingsSlice>;
}
/**
 * Visual settings model class
 */
export declare class VisualFormattingSettingsModel extends FormattingSettingsModel {
    tableSettingsCard: TableSettingsCard;
    cards: TableSettingsCard[];
}
export {};
