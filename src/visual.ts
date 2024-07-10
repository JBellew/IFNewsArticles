"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import "./../style/datatables.less"; // Import custom DataTables LESS file
import * as $ from 'jquery'; // Import jQuery first
import * as moment from 'moment';
import DOMPurify from 'dompurify';

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

import 'datatables.net';

import { VisualFormattingSettingsModel } from "./settings";

// Define the interface for Weekly News Items
interface WeeklyNewsItem {
    Title: string;
    Date: string;
    Summary: string;
    AllContent: string;
    Author: string;
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private tableElement: JQuery<HTMLElement>;

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();
        this.target = options.element;

        // Create table element
        this.tableElement = $('<div style="overflow-y: visible; overflow-x: hidden; height: 50vh;"><table id="newsTable" class="display" width="100%"></table></div>');

        // HTML for the detailed news view modal
        const MoreNewsDetailsHTML = `
        <div id="moreNews" style='display:none;' tabindex="-1" aria-labelledby="newsModalLabel" aria-hidden="true">
            <div class="news-body">
                <h3 id="newsTitle" class="newsTitle"></h3>
                <p id="newsDate" class="newsDate"></p>
                <div id="newsContent" class="newsContent"></div>
            </div>
            <div class="news-footer">
                <button type="button" id="closeNews" class="btn btn-secondary">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="currentColor" class="bi bi-x close-icon" viewBox="0 0 16 16">
                        <path d="M4.646 4.646a.5.5 0 011 0L8 6.293l2.354-2.354a.5.5 0 01.708.708L8.707 7l2.354 2.354a.5.5 0 01-.708.708L8 7.707l-2.354 2.354a.5.5 0 01-.708-.708L7.293 7 4.646 4.646a.5.5 0 010-.708z"/>
                    </svg> Close
                </button>
            </div>
        </div>
        `;

        // Append the table and modal HTML to the target element
        $(this.target).append(this.tableElement);
        $(this.target).append(MoreNewsDetailsHTML);

        // Bind close button click event to hide the modal
        $(document).on('click', '#closeNews', function() {
            $('#moreNews').hide();
        });
    }

    public update(options: VisualUpdateOptions) {
        // Populate the formatting settings model
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);

        const dataViews = options.dataViews;
        if (dataViews && dataViews[0]) {
            const dataView = dataViews[0];

            // Map dataView rows to WeeklyNewsItem objects, with default values for missing data
            const tableData = dataView.table.rows.map(row => ({
                Title: row[0] ? row[0] as string : "No Title",
                Date: row[1] ? moment(row[1] as string).format('DD/MM/YYYY') : "No Date", // Format date to Irish date format
                Summary: row[2] ? row[2] as string : "No Summary",
                AllContent: row[3] ? row[3] as string : "No Content", // Full content for the modal
                Author: row[4] ? row[4] as string : ""
            }));

            const headerColor = this.formattingSettings.tableSettingsCard.headerColor.value.value;
            const headerFontColor = this.formattingSettings.tableSettingsCard.headerFontColor.value.value;
            const tableHeight = this.formattingSettings.tableSettingsCard.tableHeight.value;
            const pageLength = this.formattingSettings.tableSettingsCard.pageLength.value;

            this.renderTable(tableData, headerColor, headerFontColor, tableHeight, pageLength);
        }
    }

    private renderTable(data: WeeklyNewsItem[], headerColor: string, headerFontColor: string, tableHeight: number, pageLength: number) {
        // Adjust table container height
        this.tableElement.css('height', `${tableHeight}vh`);

        // Destroy existing DataTable if it exists
        if ($.fn.DataTable.isDataTable(this.tableElement.find('table'))) {
            this.tableElement.find('table').DataTable().destroy();
        }

        // Clear the table content
        this.tableElement.find('table').empty();

        // Initialize DataTable with the mapped data
        const dataTable = this.tableElement.find('table').DataTable({
            data: data,
            columns: [
                { title: "Title", data: "Title" },
                { title: "Issue Date", data: "Date" },
                { title: "Summary", data: "Summary" },
                { title: "", data: "AllContent", visible: false }, // Hidden column for full content
                { title: "Author", data: "Author", visible: false } // Hidden column for author
            ],
            paging: true,
            searching: true,
            autoWidth: false,
            pageLength: pageLength, // Set page length
            order: [[1, 'desc']], // Order by the Date column (index 1) in descending order
            columnDefs: [
                { targets: 0, width: '30%' },
                { targets: 1, width: '10%' },
                { targets: 2, width: '60%' }
            ],
            language: {
                emptyTable: "No news articles to show"
            },
            dom: 'frtp' // Use default DataTables layout
        });

        // Apply header colors
        this.tableElement.find('thead th').css('background-color', headerColor);
        this.tableElement.find('thead th').css('color', headerFontColor);

        // Add click event to open news detail div
        this.tableElement.find('table tbody').on('click', 'tr', function () {
            const rowData = dataTable.row(this).data();
            $('#newsTitle').text(rowData.Title); // Insert plain text for title
            $('#newsDate').text(rowData.Date + ' ' + rowData.Author);   // Insert plain text for date and author
            $('#newsContent').html(DOMPurify.sanitize(rowData.AllContent)); // eslint-disable-line powerbi-visuals/no-implied-inner-html
            $('#moreNews').show();
            console.log(rowData.Author);
        });
    }

    public getFormattingModel() {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
