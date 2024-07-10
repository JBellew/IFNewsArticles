import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";
import "./../style/datatables.less";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import 'datatables.net';
export declare class Visual implements IVisual {
    private target;
    private formattingSettings;
    private formattingSettingsService;
    private tableElement;
    constructor(options: VisualConstructorOptions);
    update(options: VisualUpdateOptions): void;
    private renderTable;
    getFormattingModel(): powerbi.visuals.FormattingModel;
}
