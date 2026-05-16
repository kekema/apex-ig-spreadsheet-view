window.lib4x = window.lib4x || {};
window.lib4x.axt = window.lib4x.axt || {};
window.lib4x.axt.ig = window.lib4x.axt.ig || {};

/*
 * Region Plugin, enabling to add a Spreadsheet View to Interactive Grids for fast data editing, with support for copy-and-paste to and from Excel.
 * The plugin utilizes JSpreadsheet CE (https://github.com/jspreadsheet/ce).
 */
lib4x.axt.ig.spreadsheetView = (function ($) {

    const C_LIB4X_SV = 'lib4x-SV';          // has been put on the widget div (put server-side) Analog: a-IG or a-RV
    const SV_EXT = '_sv';
    const C_LIB4X_SV_EXCLUDE= 'lib4x-SV-exclude';   
    const C_SV_EDITOR = 'sv-editor';        // gets added to cells where the editor uses a popup or for additional css, so extra style or behavior
    const C_LIB4X_SV_EDITORS_CONTAINER = 'lib4x-SV-editorsContainer';   // same function as a-GV-columnItemsContainer to park editors
    const C_LIB4X_SV_TOOLBAR = 'lib4x-SV-toolbar';
    const C_LIB4X_SV_CELL_EDITOR = 'lib4x-SV-cellEditor';
    const C_IS_MAX = 'is-max';
    const C_HAS_ISSUES = 'has-issues';
    const C_APEX_THEME_VITA_DARK = 'apex-theme-vita-dark';
    const C_JSS_THEME_DARK = 'jss_theme_dark';  // css configured in jspreadsheet.themes.css
    const C_EDITOR = 'editor';
    const ORIG_GV = 'gv';
    const ORIG_SV = 'sv';  
    const OP_INSERT = 'INSERT';
    const OP_UPDATE = 'UPDATE';
    const OP_DELETE = 'DELETE';
    const FILTER_ALL = 'ALL';
    const FILTER_MODIFIED = 'MODIFIED';
    const FILTER_ISSUES = 'ISSUES';
    const TEMP_ID_PREFIX = 'svt';     
    const LOCAL_STORAGE_PREFIX = 'lib4x_ig_sv';
    let svConfig = {};                      // plugin config and options
    let sv_igStaticId = {};                 // maps the sv widget id to the IG static Id
    let ig_svStaticId = {};                 // has an entry for every IG which has a related sv
    let sv_loadAllInProgress = {};
    let sv_syncIssues = {};
    let sv_eventHandlers = {};

    let pageModule = (function() {
        /*
         * In case of starting dialogs from an inline dialog (like apex.message.confirm), no overlay is appearing to the inline dialog. 
         * The overlay is actually created (on the page body), but the z-index is lower than the inline dialog. By having next code, 
         * the z-index will be corrected and the overlay will cover the inline dialog.
         * A filterClass can be given as to restrict the check to certain dialogs only.
         */
        let enableInlineDialogOverlay = function(filterClass) 
        {
            $(apex.gPageContext$).on('dialogcreate', function(jQueryEvent, data) {
                let target$ = $(jQueryEvent.target);
                if (typeof filterClass === 'undefined' || filterClass === null || target$.closest('.ui-dialog').hasClass(filterClass))
                {
                    setTimeout(()=>{
                        if ($('.ui-widget-overlay').length > 1)
                        {
                            let maxZIndex = 0;
                            $('.ui-widget-overlay').not(":last").each(function() {
                                let zIndex = parseInt($(this).css('z-index'));
                                maxZIndex = (zIndex > maxZIndex) ? zIndex : maxZIndex;
                            });        
                            let lastZIndex = parseInt($('.ui-widget-overlay').last().css('z-index'));    
                            if (lastZIndex <= maxZIndex)
                            {
                                $('.ui-widget-overlay').last().css('z-index', maxZIndex + 1);
                                target$.dialog('moveToTop');   
                            }
                        }     
                    }, 10);
                }
            });             
        }   
        enableInlineDialogOverlay(null);
    })();

    let gridModule = (function () {
        /*
         * Init/Adjust the IG's for which a SV is configured. 
         */
        let initIGs = function (filterIgStaticId, filterClass, svStaticId) {
            let svStaticIdSv = svStaticId + SV_EXT;
            let config = svConfig[svStaticIdSv];
            apex.gPageContext$.on("apexreadyend", function (jQueryEvent) {
                if (filterIgStaticId)
                {
                    let igRegion = apex.region(filterIgStaticId);
                    // test if region is an IG
                    if (igRegion?.type != 'InteractiveGrid') {
                        throw new Error('IG Spreadsheet View Settings error: \'' + filterIgStaticId + '\' is not an Interactive Grid Region');
                    }
                    // safety check: gridView should be enabled
                    if (!igRegion.widget().interactiveGrid('option').config?.views?.grid?.features?.gridView) {
                        throw new Error('IG Spreadsheet View can not be attached to IG \'' + filterIgStaticId + '\': the IG gridView feature is not enabled');
                    }                    
                }
                // determine applicable IG region(s)
                let igRegions = Object.values(apex.regions).filter(r => {
                    if (r.type !== "InteractiveGrid") return false;
                    if (r.element.hasClass(C_LIB4X_SV_EXCLUDE)) return false;
                    if (filterIgStaticId) {
                        return r.element.attr('id') === filterIgStaticId;
                    }
                    if (filterClass) {
                        return r.element.hasClass(filterClass);
                    }                    
                    return true;
                });
                // add Spreadsheet View button to each
                igRegions.forEach((igRegion) => {
                    // only attach to IG's with gridView enabled
                    if (igRegion.widget().interactiveGrid('option').config?.views?.grid?.features?.gridView) 
                    {
                        const igStaticId = igRegion.element.attr('id'); // use const as we refer to igStaticId in event handler (fresh copy)
                        // skip if the IG already has a SV enabled
                        if (!ig_svStaticId.hasOwnProperty(igStaticId))
                        {
                            ig_svStaticId[igStaticId] = svStaticId;
                            // add Spreadsheet View button to IG toolbar - create supporting action first              
                            const igActions = igRegion.call('getActions');
                            let model = igRegion.call('getViews').grid.model;
                            // when the IG started in another view like detail view, or the IG is hidden (like in a tab),
                            // the IG columns won't be available yet. In that case, create with disabled button
                            let disableSVOpen = (igRegion.call('getViews').grid.view$?.grid('getColumns') == null);
                            igActions.add({
                                name: 'lib4x-sv-open',
                                icon: 'fa fa-table-file',
                                title: getMessage('SPREADSHEET'),
                                disabled: disableSVOpen,
                                action: function (event) {
                                    spreadsheetViewModule.openSpreadsheetView(svStaticId, igStaticId);
                                }
                            });
                            // optionally, add a 'Revert All'
                            if (config.options.buttons?.ig?.revertAll && model.getOption('editable'))
                            {
                                igActions.add({                                
                                    name: 'lib4x-revert-all',
                                    label: getMessage('REVERT_ALL'),
                                    icon: 'fa fa-undo',
                                    title: getMessage('REVERT_ALL_CHANGES'),
                                    action: function (event) {
                                        revertAll(igStaticId);
                                    }
                                });
                            }
                            if (disableSVOpen) {
                                apex.widget.util.onVisibilityChange(apex.region(igStaticId).widget(), function(visible){
                                    if (visible && (igRegion.call('getCurrentViewId') == 'grid'))
                                    {
                                        igActions.enable('lib4x-sv-open'); 
                                    }                                 
                                });              
                            }
                            // buttons only enabled in case of IG grid view
                            $('#' + igStaticId).on("interactivegridviewchange", function (jQueryEvent, data) {
                                if (data?.view == 'grid') {
                                    igActions.enable('lib4x-sv-open');
                                    if (config.options.buttons?.ig?.revertAll)
                                    {
                                        igActions.enable('lib4x-revert-all');
                                    }
                                }
                                else
                                {
                                    igActions.disable('lib4x-sv-open');
                                    if (config.options.buttons?.ig?.revertAll)
                                    {
                                        igActions.disable('lib4x-revert-all');
                                    }
                                }
                            });                            
                            // add button to IG toolbar for opening the SV
                            let tbWidget = igRegion.widget().interactiveGrid('getToolbar');
                            let actionsMenuIsHidden = tbWidget.find('button[id$="actions_button"]').is(':hidden');
                            let doRefresh = false;
                            let tbViewsGroup = tbWidget.toolbar('findGroup', 'views');
                            // tbViewsGroup is always present also when detail and icon views are switched off
                            // still test for presence just in case
                            if (tbViewsGroup?.controls && Array.isArray(tbViewsGroup?.controls)) 
                            {
                                let actionsMenuIsHidden = tbWidget.find('button[id$="actions_button"]').is(':hidden');
                                tbViewsGroup.controls.push({ type: 'BUTTON', action: 'lib4x-sv-open', iconOnly: true });
                                doRefresh = true;
                            }
                            // add Revert All button
                            if (config.options.buttons?.ig?.revertAll && model.getOption('editable'))
                            {
                                let tbActions3Group = tbWidget.toolbar('findGroup', 'actions3');
                                if (tbActions3Group?.controls && Array.isArray(tbActions3Group?.controls)) 
                                {
                                    tbActions3Group.controls.push({ type: 'BUTTON', action: 'lib4x-revert-all', iconBeforeLabel: true });
                                    doRefresh = true;
                                }
                            }
                            if (doRefresh)  
                            {
                                tbWidget.toolbar('refresh');
                                if (actionsMenuIsHidden)
                                {
                                    // a toolbar refresh is restoring all toolbar items and then hides any items as per (action) settings
                                    // however if the actions menu is configured to not to be shown, it fails to hide this item, so we do it:
                                    tbWidget.find('button[id$="actions_button"]').hide();
                                }   
                                // APEX adds a 'a-Button--withIcon' class which reduces the padding but the calculation ends up with
                                // a button height less as compared to buttons with text only - doesn't look good, so we remove the class
                                // tested with button border width 2px, 3px: still looks good
                                $('#' + igStaticId).find('button[data-action="lib4x-sv-open"]').removeClass('a-Button--withIcon');                                                             
                            }                
                        }
                    }
                });
            });

            // revert all model changes: inserts, updates and deletes
            function revertAll(igStaticId)
            {
                let gridView = apex.region(igStaticId).call('getViews').grid;
                let model = gridView.model;
                let changes = model.getChanges();
                if (changes.length > 0)
                {
                    function revertChanges()
                    {
                        model.revertRecords(changes.filter((recMeta)=>recMeta.updated||recMeta.deleted).map((recMeta) => recMeta.record));
                        model.deleteRecords(changes.filter((recMeta)=>recMeta.inserted).map((recMeta) => recMeta.record));  
                    }
                    // if multiple changes, ask confirmation
                    if (changes.length > 1)
                    {
                        apex.message.confirm(getMessage('REVERT_ALL_CHANGES') + '? (' + changes.length + ' rows)', function(okPressed) {
                            if (okPressed) 
                            {
                                revertChanges();
                            }
                        }, {iconClasses: 'fa fa-undo-alt fa-2x'});
                    }
                    else
                    {
                        revertChanges();
                    }
                }
                // exit any edit mode
                gridView.curInst.inEditMode() ? gridView.curInst.setEditMode(false) : null;
            }
        }

        return {
            initIGs: initIGs
        }
    })();

    let spreadsheetViewModule = (function () {
        let tempRecordSeq = 0;                  // facilitates temp record id's for inserted but not yet synchronized rows
        let currentPageNumber = 0;              // in case of page pagination, keeps track of the current page
        let lastClosedIgStaticId = null;        // static id of IG as related to the SV at moment of last close
        let localStorage = null;

        if (apex.storage.hasLocalStorageSupport()) {
            localStorage = apex.storage.getScopedLocalStorage({prefix: LOCAL_STORAGE_PREFIX, useAppId: false, usePageId: false});
        }

        /*
         * Create a Spreadsheet View based on the IG (columns) definition and model data.
         * svStaticId: the inline dialog region id
         * isRecreate: true in case an existing view is recreated as because of synchronize or load all
         */
        function createSpreadsheetView(svStaticId, isRecreate) {
            let svStaticIdSv = svStaticId + SV_EXT;             // widget id
            let igStaticId = sv_igStaticId[svStaticIdSv];
            let igRegion = apex.region(igStaticId);
            let gridView = igRegion.call('getViews').grid;
            let config = svConfig[svStaticIdSv];
            tempRecordSeq = 0;
            let actionsContext = apex.actions.findContext('IGSpreadsheetView', $('#' + svStaticIdSv)[0]);
            let ctxPrototype = null;

            /*
             * Create an instance of the JSS Spreadsheet with one worksheet.
             * wsData: JSS worksheet data
             * wsIds: record id's as will be attached to the rows elements
             * wsMeta: JSS worksheet metadata. Currently not used: to restrictive datastructure and also issues when inserting/deleting rows
             * dsMeta: dataset metadata, comparable to the IG record metadata
             * wsReadOnlyCells: cells to be marked as read only
             * footers: for aggregates
             * nestedHeaders: used for column group headers
             * hasHighlights: true if IG has highlighting applied. If true, a toggle button to show/hide highlighting will be added
             */
            function createSpreadsheet(svColumns, numberOfRows, wsData, wsIds, wsMeta, dsMeta, wsReadOnlyCells, svAggregators, footers, nestedHeaders, hasHighlights) {
                let dlg$ = $('#' + svStaticId);
                let body$ = dlg$.find('.t-DialogRegion-body, .t-DrawerRegion-body');
                let bodyWrapperOut$ = dlg$.find('.t-DialogRegion-bodyWrapperOut, .t-DrawerRegion-bodyWrapperOut');
                let hPadding = body$.innerHeight() - body$.height();
                let wPadding = body$.innerWidth() - body$.width();

                let setDialogCoordinates = true;
                let tableWidth = null;
                let saveScrollTop = null;
                let saveScrollLeft = null;
                let savePageNumber = 0;
                // destroy existing instance in case the spreadsheet is recreated like in case of Load All, Synchronize            
                if (getJSpreadsheet(svStaticIdSv))
                {  
                    let jssContent$ = dlg$.find('.jss_content');
                    tableWidth = jssContent$.outerWidth();     // in case of recreate, take current width to prevent 'flickering'  
                    saveScrollTop = jssContent$.scrollTop();
                    saveScrollLeft = jssContent$.scrollLeft();
                    savePageNumber = currentPageNumber;
                    destroySpreadsheet(svStaticId);
                    setDialogCoordinates = false;
                }
                let toolbarHeight = 43;     // assume 43 as long as no toolbar instance created yet
                currentPageNumber = 0;
                let lastEventWasMouse = false;      // used to see if a cell selection was from mousedown 
                let lastInteractionEvent = null;
                // determine worksheet options
                let paginationType = config.paginationType;
                if (paginationType == 'PT_IG_SETTING')
                {
                    let igPagination = apex.region(igStaticId).call('getViews').grid.view$.grid('option', 'pagination');
                    paginationType = igPagination?.scroll ? 'PT_SCROLL' : 'PT_PAGE';
                }     
                let {worksheet: worksheetOptions = {}, ...spreadsheetOptions} = config.options?.jspreadsheetOptions ?? {};
                if (paginationType == 'PT_PAGE')
                {
                    if (!worksheetOptions.pagination)
                    {
                        worksheetOptions.pagination = 10;
                    }
                    if (!worksheetOptions.paginationOptions)
                    {
                        worksheetOptions.paginationOptions = [10, 15, 20, 25, 50, 100];
                    }                    
                }
                else
                {
                    worksheetOptions.pagination = null;
                }
                if (worksheetOptions?.pagination)
                {
                    // pagination page size is kept in local storage
                    if (localStorage) {
                        let pageSize = localStorage.getItem('paginationPageSize');
                        if (pageSize)
                        {
                            worksheetOptions.pagination = Number(pageSize);
                        }
                    }                    
                }
                // regarding width/height, we take a first shot
                // upon sv visibility change to visible, a dialogresize is triggered 
                // which will size the worksheet more accurate
                let paginationbarHeight = worksheetOptions?.pagination ? 56 : 0;
                let bodyHeight = bodyWrapperOut$.height() ?? 0;
                let tableHeight = bodyHeight - hPadding - toolbarHeight - paginationbarHeight - 3;
                let worksheetName = getWorksheetName(igStaticId);
                // instantiate JSS; always one worksheet only
                let worksheets = jspreadsheet($('#' + svStaticIdSv)[0], {
                    ...spreadsheetOptions,
                    tabs: false,
                    toolbar: false,           // no JSS toolbar, we create our own APEX style toolbar      
                    worksheets: [
                        {
                            ...worksheetOptions,
                            worksheetName: worksheetName,           // not really required but setting the worksheet name - same as dialog title
                            minDimensions: [svColumns.length, 0],   // rows as 0: it will just take as per the number of rows in wsData
                            tableOverflow: true,
                            tableHeight: tableHeight + "px",
                            tableWidth: tableWidth ? tableWidth + "px" : tableWidth,
                            columns: svColumns,
                            allowInsertColumn: false,
                            allowManualInsertColumn: false,
                            allowDeleteColumn: false,
                            allowRenameColumn: false,
                            columnDrag: false,                      // stick to IG column order
                            rowDrag: false,
                            data: wsData,
                            meta: wsMeta,
                            footers: footers,
                            search: true,                           // use search input which we will merge into the toolbar
                            filters: false,                         // has several issues 
                            nestedHeaders: nestedHeaders                         
                        }
                    ],
                    // suppress all context menu's - all functions will be accessible from the toolbar
                    contextMenu: function () {
                        return false;
                    },
                    onbeforeselection: function (instance, x1, y1, x2, y2) 
                    {
                        if (y1 == null)
                        {
                            // illegal selection - prevent further js error in JSS      
                            // this can happen as because of a small bug in JSS where in case you first click in the worksheet margin and
                            // then make a row selection using shift, it throws a can not read property because of undefined error
                            // this situation is also further prevented by not letting mousedown clicks on the toolbar bubble to the worksheet
                            return false;
                        }
                    },
                    onselection: function (instance, x1, y1, x2, y2) 
                    {      
                        // upon selecting a row or column, JSS is not closing any open editor
                        if (instance.selectedRow || instance.selectedHeader)
                        {
                            closeAnyEditor(instance);
                        }
                        // in case of edit on focus, open the editor if a single cell is selected
                        if (actionsContext.lookup('toggle-edit-on-focus').editOnFocus)
                        {
                            if (x1 === x2 && y1 === y2) {
                                // exclude lib4 simple choice columns as it would immediately toggle the choice state
                                let svColumn = instance.options.columns[x1];
                                let columnType = svColumn.lib4x?.columnType;
                                if (!['lib4x_checkbox', 'lib4x_simple_radio', 'lib4x_switch', 'lib4x_pill_buttons'].includes(columnType))
                                {
                                    let cell = instance.getCell(x1, y1);
                                    let lastEvent = lastInteractionEvent;
                                    setTimeout(() => {
                                        // check, double check if not (a) cell in edit mode currently
                                        if (instance.edition == null)
                                        {                                        
                                            if (!$(cell).hasClass(C_EDITOR))
                                            {      
                                                instance.openEditor(cell, false, lastEvent);
                                            }
                                            else
                                            {
                                                apex.debug.info('lib4x-SV: open editior upon edit on focus: cell has editor already open', x1, y1);
                                            }
                                        }
                                    }, 0);
                                }
                            }
                        }
                        // keep track of last selection so we can restore it if needed when the selection is gone upon undo, redo
                        // also used when detecting for any open editor cell
                        instance.options.lib4x.lastSelection = instance.getSelection();
                    },                   
                    oncopy: function(instance, selectedRange, copiedData, cut)
                    {
                        // prepare clipboard content, both text and html versions
                        if (selectedRange && selectedRange.length == 4)
                        {
                            // htmlBuilder will escape html element content
                            let out = apex.util.htmlBuilder();
                            let textResult = '';
                            // next is already added by the clipboard API: '<html><body><!--StartFragment-->'
                            // Future: enable extra style from configuration? (also on TD level?)
                            out.markup('<style>table {mso-displayed-decimal-separator: "."; mso-displayed-thousand-separator: ","; white-space: nowrap;}</style>');
                            out.markup('<table><tbody>');
                            let svColumns = instance.options.columns;
                            let lastTextValue = '';  
                            // render table of values                          
                            for (let y = selectedRange[1]; y <= selectedRange[3]; y++)
                            {
                                let filtered = instance.results;
                                if (!filtered || filtered.includes(y))
                                {
                                    out.markup('<tr>');
                                    let sep = '';
                                    for (let x = selectedRange[0]; x <= selectedRange[2]; x++)
                                    {
                                        let svColumn = svColumns[x];
                                        let cbValue;
                                        if (svColumn.lib4x?.toClipboardValue)
                                        {
                                            cbValue = svColumn.lib4x?.toClipboardValue(instance, x, y);
                                        } 
                                        else
                                        {
                                            cbValue = {textValue: '', htmlValue: ''};
                                            let dataValue = instance.getValueFromCoords(x, y, false);   
                                            dataValue = util.getDisplayValue(dataValue);                                     
                                            if (dataValue)
                                            {
                                                cbValue = {textValue: dataValue, htmlValue: dataValue};
                                            }
                                        }                            
                                        textResult = textResult + sep + cbValue.textValue;
                                        lastTextValue = cbValue.textValue;
                                        let tdStyle = '';
                                        if (cbValue.msoNumberFormat)
                                        {
                                            tdStyle = ` style="mso-number-format:'${cbValue.msoNumberFormat}';"`;
                                        }                              
                                        out.markup('<td' + tdStyle + '>').content(cbValue.htmlValue).markup('</td>');                                
                                        sep = '\t';
                                    }
                                    textResult = textResult + '\n';
                                    out.markup('</tr>'); 
                                }                           
                            }
                            if ((lastTextValue === '') && ((selectedRange[1] == selectedRange[3]) || (selectedRange[0] == selectedRange[2])))
                            {
                                // below extra \n should not be needed but without, JSS won't see when the last cell is empty upon paste 
                                // even resulting in console error when one empty cell is copied
                                // it's only when selecting cell(s) in one row or one column
                                textResult = textResult + '\n';
                            }
                            out.markup('</tbody></table>');
                            // next is added by the clipboard API: '<!--EndFragment--></body></html>'
                            let htmlResult = out.toString();
                            // write result to the clipboard. Effectively, the html will have the html element twice, which doesn't seem an issue though
                            setTimeout(() => {
                                navigator.clipboard.write([
                                    new ClipboardItem({
                                        'text/html': new Blob([htmlResult], { type: 'text/html' }),
                                        'text/plain': new Blob([textResult], { type: 'text/plain' })
                                    })
                                ]).catch((primaryError) => {
                                    return navigator.clipboard.writeText(textResult)
                                        .catch((fallbackError) => {
                                            fallbackError.primaryError = primaryError;
                                            throw fallbackError;
                                        });
                                }).catch((error) => {
                                    apex.debug.info('Not able to write to the Clipboard.', {
                                        error,
                                        primaryError: error.primaryError
                                    });
                                });
                            }, 0);
                        }
                        return false;
                    },
                    onbeforechange(instance, cell, colIndex, rowIndex, newValue)
                    {
                        if (!instance.options.lib4x.suppressChangeEvent)
                        {
                            let svColumn = instance.options.columns[colIndex];
                            // In theory, values can contain CR/LF (\r\n) for example if the value comes via copy/paste from Windows. 
                            // APEX uses \n only (unix standard) so we convert here any occurences 
                            if (newValue && util.valueIsString(newValue)) {  
                                newValue = newValue.replace(/\r\n/g, '\n');
                            }
                            let oldValue = instance.getValueFromCoords(colIndex, rowIndex, false);
                            let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);
                            if (!util.equalValues(newValue, oldValue))
                            {
                                if (!dsMetaUtil.fieldMetaHasProperty(rowMeta, ORIG_SV, svColumn.name, 'origValue'))
                                {
                                    dsMetaUtil.setFieldMeta(rowMeta, ORIG_SV, svColumn.name, 'origValue', oldValue);
                                }
                            }
                            // preserve oldValue so it can be used in onchange event handler
                            dsMetaUtil.setFieldMeta(rowMeta, ORIG_SV, svColumn.name, 'oldValue', oldValue);     
                        }                       
                        return newValue;
                    },
                    onchange(instance, cell, colIndex, rowIndex, newValue) 
                    {
                        // process change into dataset metadata and change markup classes
                        if (!instance.options.lib4x.suppressChangeEvent)
                        {                        
                            let svColumn = instance.options.columns[colIndex];                            
                            let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);  
                            let oldValue = rowMeta.sv.fields[svColumn.name].oldValue;   
                            // when same value is pasted, still onchange is being called by JSS
                            if (!util.equalValues(newValue, oldValue))
                            {                                      
                                updateCellChangeMarking(instance, colIndex, rowIndex, newValue);
                                actionsContext.disable('load-all');
                                actionsContext.disable('switch-pagination-type');
                                // fire SV event
                                let ctx = Object.create(ctxPrototype);
                                ctx.rowIndex = Number(rowIndex);
                                ctx.columnName = svColumn.name;
                                ctx.oldValue = toCanonicalValue(svColumn, oldValue);
                                ctx.newValue = toCanonicalValue(svColumn, newValue);                            
                                setTimeout(() => {
                                    fireSVEvent(igStaticId, 'onChange', ctx, instance, true);
                                });                              
                            }
                        }
                    },
                    onafterchanges(instance, changes)
                    {
                        instance.options.lib4x.aggrEngine?.applyChanges(changes, 'set');
                    },
                    onundo(instance, info)
                    {
                        // process undo wrt dataset metadata and change markup classes
                        if (info?.action == 'setValue')
                        {
                            instance.options.lib4x.aggrEngine?.applyChanges(info.records, 'undo');
                            info.records.forEach((record, index) => {
                                let currentValue = instance.getValueFromCoords(record.col, record.row, false);
                                updateCellChangeMarking(instance, record.col, record.row, currentValue);
                                let svColumn = instance.options.columns[record.col];
                                let ctx = Object.create(ctxPrototype);
                                ctx.rowIndex = Number(record.row);
                                ctx.columnName = svColumn.name;
                                ctx.oldValue = toCanonicalValue(svColumn, record.value);
                                ctx.newValue = toCanonicalValue(svColumn, record.oldValue);  // with undo, newValue becomes the record.oldValue
                                setTimeout(() => {
                                    fireSVEvent(igStaticId, 'onChange', ctx, instance, true);
                                });                              
                            });
                        }
                        else if (info?.action == 'insertRow')
                        {
                            info.rowNode.forEach((row) => {
                                let recordId = $(row.element).data('id');
                                let rowMeta = dsMetaUtil.getRecordMeta(instance, recordId);
                                rowMeta.sv.transient = true;
                            });
                            instance.options.lib4x.aggrEngine?.calculateAll();
                        }
                        else if (info?.action == 'deleteRow')
                        {
                            info.rowNode.forEach((row) => {
                                let recordId = $(row.element).data('id');
                                let rowMeta = dsMetaUtil.getRecordMeta(instance, recordId);
                                delete rowMeta.sv.deleted;
                                delete rowMeta.sv.transient;
                                // upon a delete action, JSS is removing the lib4x tr classes, so we restore those here
                                if (rowMeta.gv.inserted || rowMeta.sv.inserted)
                                {
                                    setRowInsertedClass(instance, row.y, rowMeta);
                                }
                                if (rowMeta.gv.updated || rowMeta.sv.updated)
                                {
                                    setRowUpdatedClass(instance, row.y, rowMeta);
                                }
                            });
                            instance.options.lib4x.aggrEngine?.calculateAll();
                        }  
                        // check if selection to be restored
                        if (instance.options.lib4x.lastSelection)
                        {
                            if (!instance.getSelection())
                            {
                                // selection is gone so restore it
                                let selection = instance.options.lib4x.lastSelection;
                                instance.updateSelectionFromCoords(selection[0], selection[1], selection[2], selection[3]);
                            }
                        }      
                        if (!dataHasChanged(svStaticId))
                        {
                            actionsContext.enable('load-all');
                            actionsContext.enable('switch-pagination-type');
                        }                      
                    },   
                    onredo(instance, info)
                    {
                        // process redo wrt dataset metadata and change markup classes
                        if (info?.action == 'setValue')
                        {
                            instance.options.lib4x.aggrEngine?.applyChanges(info.records, 'redo');
                            info.records.forEach((record, index) => {
                                let currentValue = instance.getValueFromCoords(record.col, record.row, false);
                                updateCellChangeMarking(instance, record.col, record.row, currentValue);  
                                let svColumn = instance.options.columns[record.col]; 
                                let ctx = Object.create(ctxPrototype);        
                                ctx.rowIndex = Number(record.row);
                                ctx.columnName = svColumn.name;
                                ctx.oldValue = toCanonicalValue(svColumn, record.oldValue);
                                ctx.newValue = toCanonicalValue(svColumn, record.value);
                                setTimeout(() => {
                                    fireSVEvent(igStaticId, 'onChange', ctx, instance, true);
                                });                                  
                            });
                        }
                        else if (info?.action == 'insertRow')
                        {
                            info.rowNode.forEach((row) => {
                                let recordId = $(row.element).data('id');
                                let rowMeta = dsMetaUtil.getRecordMeta(instance, recordId);
                                delete rowMeta.sv.transient;         
                            });
                        }                         
                        else if (info?.action == 'deleteRow')
                        {
                            info.rowNode.forEach((row) => {
                                let recordId = $(row.element).data('id');
                                let rowMeta = dsMetaUtil.getRecordMeta(instance, recordId);
                                if (rowMeta.sv.inserted)
                                {
                                    rowMeta.sv.transient = true;
                                }
                                else
                                {
                                    rowMeta.sv.deleted = true;
                                }
                            });
                        }
                        // check if selection to be restored
                        if (instance.options.lib4x.lastSelection)
                        {
                            if (!instance.getSelection())
                            {
                                // selection is gone so restore it
                                let selection = instance.options.lib4x.lastSelection;
                                instance.updateSelectionFromCoords(selection[0], selection[1], selection[2], selection[3]);
                            }
                        }                                          
                        if (dataHasChanged(svStaticId))
                        {
                            actionsContext.disable('load-all');
                            actionsContext.disable('switch-pagination-type');
                        }                                                 
                    },
                    onbeforepaste: function (instance, data, colIndex, rowIndex) {
                        // conversion and normalization of clipboard data
                        // notice: when pasting a value for a (text) cell which is equal to the 
                        // current value, still onBeforeChange/onChange will be called by JSS
                        // which is not causing any issue, but it gets added to the history
                        if (!Array.isArray(data)) return data;                        
                        
                        function extractSeparators(doc) 
                        {
                            function unescapeSeparator(value) 
                            {
                                // "\," → ","
                                // "\." → "."
                                return value.replace(/^\\/, '');
                            } 
                            let decimal = null;
                            let thousands = null;

                            doc.querySelectorAll('style').forEach(styleEl => {
                                let css = styleEl.textContent;
                                // remove any HTML comment wrappers (as put by Excel)
                                css = css.replace(/<!--|-->/g, '');
                                const decMatch = css.match(/mso-displayed-decimal-separator\s*:\s*"([^"]+)"/i);
                                const thouMatch = css.match(/mso-displayed-thousand-separator\s*:\s*"([^"]+)"/i);
                                if (decMatch) 
                                {
                                    decimal = unescapeSeparator(decMatch[1]);
                                }
                                if (thouMatch) 
                                {
                                    thousands = unescapeSeparator(thouMatch[1]);
                                }
                            });
                            return {decimal, thousands};
                        }

                        function parseForNumber(text, decimalSep, thousandSep) 
                        {
                            if (!text) return '';
                            const isNegative = /\(.*\)/.test(text) || text.includes('-');
                            let clean = text.replace(/[^\d.,]/g, '');
                            let normalized = clean.split(thousandSep).join('').replace(decimalSep, '.');
                            let value = parseFloat(normalized);
                            if (isNaN(value)) return text;
                            if (isNegative) value *= -1;
                            if (text.includes('%')) value /= 100;
                            return String(value);
                        }                        

                        // transform any numbers as to comply with APEX app locale settting for decimal separator
                        let cbSeparators = null;
                        let html = instance.options.lib4x.clipboardHtml;
                        delete instance.options.lib4x.clipboardHtml;    // removes the property
                        if (html)
                        {
                            try
                            {
                                let doc = new DOMParser().parseFromString(html, 'text/html');
                                cbSeparators = extractSeparators(doc);
                            }
                            catch(e){};
                        }
                        // iterate data rows/columns
                        const normalized = data.map(row => {
                            let x = Number(colIndex);
                            return row.map(cell => {
                                let svColumn = instance.options.columns[x];
                                x = x + 1;
                                if (cell && typeof cell === 'object' && 'value' in cell) 
                                {
                                    let value = cell.value;
                                    if (value && svColumn)
                                    {
                                        if (svColumn.lib4x?.columnType == 'lib4x_number' && cbSeparators && cbSeparators.decimal && cbSeparators.thousands)
                                        {
                                           value = parseForNumber(value, cbSeparators.decimal, cbSeparators.thousands);
                                        }
                                        else if (svColumn.lib4x?.columnType == 'lib4x_date_picker')
                                        {
                                            // copy/paste dates between Excel and SV remains a bit tricky
                                            // in most cases, dates do come fine, but if not, the onBeforePasteDate hook can be utilized to get it right
                                            if (config.options.onBeforePasteDate)
                                            {
                                                let eventObject = {columnName: svColumn.name, targetFormat: svColumn.customFormat, value: value};
                                                config.options.onBeforePasteDate(eventObject);
                                                value = eventObject.value;    
                                            }
                                        }
                                        if (svColumn.lib4x?.inputToDataValue)
                                        {
                                            // process input value to data value
                                            value = svColumn.lib4x.inputToDataValue(value);
                                        }    
                                    }
                                    return value; 
                                }
                                return cell; // already a string or null
                            })
                        });
                        return normalized;
                    },
                    // for any future need
                    /*onpaste: function(instance, pastedInfo)
                    {

                    },*/
                    onbeforeinsertrow: function(instance, rows)
                    {
                        return gridView.model.allowAdd() && (instance.getData().length + rows?.length <= (config.options.maxRows + config.options.maxAdditionalRows));
                    },                    
                    oninsertrow: function(instance, rows)
                    {
                        // set any default values
                        // couldn't get setting defaults to work in onbeforeinsertrow event or insertRow method, so doing it here
                        instance.ignoreHistory = true;
                        instance.options.lib4x.suppressChangeEvent = true;
                        let svColumns = instance.options.columns;
                        rows.forEach((row) => { 
                            let recordId = getTempRecordId(); 
                            $(instance.rows[row.row].element).data('id', recordId);
                            let rowMeta = dsMetaUtil.initRowMeta();
                            rowMeta.sv.inserted = true;
                            instance.options.lib4x.dsMeta.set(recordId, rowMeta);      
                            svColumns.forEach((svColumn, colIndex) => {
                                let svValue = svColumn.lib4x.defaultValue || '';
                                if (svColumn.lib4x.modelToDataValue)
                                {
                                    svValue = svColumn.lib4x.modelToDataValue(svValue);
                                } 
                                instance.setValueFromCoords(colIndex, row.row, svValue, true);
                            });               
                        });
                        instance.ignoreHistory = false;  
                        instance.options.lib4x.suppressChangeEvent = false;
                        updateRowInsertMarking(instance, rows);
                        instance.options.lib4x.aggrEngine?.calculateAll();
                        actionsContext.disable('load-all');
                        actionsContext.disable('switch-pagination-type');
                    },
                    onbeforedeleterow: function(instance, rows)
                    {
                        // check if all rows can be deleted
                        //
                        // backup current selection setting
                        instance.options.lib4x.selectionBefore = instance.getSelection();
                        let igConfig = igRegion.call('option').config;
                        if (!(igConfig.editable && igConfig.editable.allowedOperations?.delete))
                        {
                            return false;
                        }
                        let rowsToDelete = [];
                        let allCanBeDeleted = true;
                        instance.options.lib4x.onDeleteRows.clear();
                        rows.forEach((rowIndex) => {
                            let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);                           
                            let deleteAllowed = rowMeta.gv.deleteAllowed;
                            // deleteAllowed meta property might be absent and so undefined, so explicitely test for false
                            // if absent, allow delete
                            if (!(deleteAllowed === false))
                            {
                                rowsToDelete.push(rowIndex);                             
                            }
                            else
                            {
                                allCanBeDeleted = false;
                            }
                            // track the rowMeta's so we can update the rowMeta in ondeleterow event
                            instance.options.lib4x.onDeleteRows.set(rowIndex, rowMeta);
                        });
                        // returning rowsToDelete one should expect to work (see JSS doc), but it doesn't
                        // so we can only just return false if any one can not be deleted
                        if (!allCanBeDeleted)
                        {
                            // No rows deleted - not all selected rows are allowed to be deleted
                            apex.message.alert(getMessage('NO_ROWS_DELETED_NOT_ALL_ALLOWED'));
                            return false;
                        }
                    },
                    ondeleterow: function(instance, rows)
                    {
                        // update dataset metadata (deleted rows are continued to being kept in ds metadata)
                        let lastRowMeta = null;
                        if (instance.rows.length == 1)
                        {
                            // JSS won't allow to delete the last row. So if this row was 
                            // included in the the rows to delete, skip this one when updating rowMeta.
                            lastRowMeta = dsMetaUtil.getRowMeta(instance, 0);
                        }                          
                        rows.forEach((rowIndex) => {
                            let rowMeta = instance.options.lib4x.onDeleteRows.get(rowIndex);  
                            if (rowMeta !== lastRowMeta)
                            {                     
                                if (rowMeta.sv.inserted)
                                {
                                    rowMeta.sv.transient = true;
                                }
                                else
                                {
                                    rowMeta.sv.deleted = true;
                                }
                            }
                            else
                            {
                                // It is not possible to delete the last row
                                apex.message.alert(getMessage('DELETE_LAST_ROW_NOT_POSSIBLE'));
                            }
                        });
                        // if selection is gone, restore it
                        if (instance.options.lib4x.selectionBefore)
                        {
                            if (!instance.getSelection())
                            {
                                // selection is gone so restore it
                                let selection = instance.options.lib4x.selectionBefore;
                                instance.updateSelectionFromCoords(selection[0], selection[1], selection[2], selection[3]);
                                instance.options.lib4x.selectionBefore = null;
                            }                           
                        }
                        instance.options.lib4x.onDeleteRows.clear();
                        instance.options.lib4x.aggrEngine?.calculateAll();    
                        actionsContext.disable('load-all'); 
                        actionsContext.disable('switch-pagination-type');                 
                    },
                    //oncreateeditor: function(instance, cell, x, y, input, options) 
                    //{
                    //    let svColumn = instance.options.columns[x];                     
                    //},
                    oneditionstart: function(instance, cell, x, y, a, b, c)
                    {     
                        // any (input) element/value preparations
                        let svColumn = instance.options.columns[x];
                        if (svColumn.type == 'text')
                        {
                            setTimeout(() => {
                                let input = cell.querySelector('input');
                                if (svColumn.lib4x?.maxlength)
                                {              
                                    // maxlength can be set in page designer. APEX is not including it in server-side validation                
                                    $(input).attr('maxlength', svColumn.lib4x.maxlength);
                                    // Future: see if anyhow also custom attributes can be included as they are defined in page designer
                                }
                                if (svColumn.lib4x?.textCase)
                                {             
                                    // this will only be effective upon typing
                                    // see onbeforechange/inputtodatavalue for transforming the data value 
                                    $(input).attr('data-text-case', svColumn.lib4x.textCase);
                                }                       
                            });
                        }  
                        if (!lastEventWasMouse && actionsContext.lookup('toggle-edit-on-focus').editOnFocus)                     
                        {
                            if (!(lastInteractionEvent?.key == 'F2'))
                            {
                                setTimeout(()=>{
                                    let inputElement = $(cell).find('input')[0];
                                    if (inputElement && inputElement.tagName === 'INPUT' && inputElement.type == 'text') 
                                    {                                      
                                        inputElement.select();
                                    }
                                });
                            }
                        }                     
                        else if (lastEventWasMouse && lastInteractionEvent)
                        {
                            if (svColumn.align == 'left')
                            {
                                let lastEvent = lastInteractionEvent;                                  
                                setTimeout(()=>{
                                    let inputElement = $(cell).find('input')[0];
                                    if (inputElement && inputElement.tagName === 'INPUT' && inputElement.type == 'text') 
                                    {                                     
                                        let pos = document.caretPositionFromPoint(lastEvent.clientX, lastEvent.clientY);
                                        if (pos && pos.offset != null) 
                                        {
                                            inputElement.setSelectionRange(pos.offset, pos.offset);
                                        }          
                                    }                                              
                                }, 50);
                            }                            
                        }
                        lastEventWasMouse = false;
                        lastInteractionEvent = null;                                    
                    },
                    onchangepage: function(instance, pageNumber) {
                        // for page pagination, keep track of current page number
                        // so we can keep the current page in case of a SV recreate
                        currentPageNumber = pageNumber;
                    },                  
                    oneditionend: function(instance, cell, x, y, v, save)
                    {
                        // because of a quirk in JSS, after editing, the focus is not returned to the spreadsheet container
                        // which will cause 'escape' key to be unresponsive
                        // so we set the focus ourselves
                        setTimeout(function () {
                            $('#' + svStaticIdSv).focus();
                        });                        
                    },   
                    // generic JSS event handler
                    // any future use                 
                    onevent: function(eventname, instance, ...args)
                    {
                        //console.log('Event: ', eventname, args);
                    }
                });
                let instance = worksheets[0];     
                // single cell select with mouse can trigger edit on focus  
                $(instance.element).on('mousedown.lib4x_jss', function(jQueryEvent, data) {
                    lastEventWasMouse = true;
                    lastInteractionEvent = jQueryEvent.originalEvent;
                });
                $('#'+svStaticIdSv).off('keydown.lib4x_jss').on('keydown.lib4x_jss', function(jQueryEvent, data) {
                    lastEventWasMouse = false;
                    lastInteractionEvent = jQueryEvent.originalEvent;
                });               
                // set the record id's on the row elements
                // via the record id, we can get the row metadata                       
                instance.rows.forEach((row) => {
                    $(row.element).data('id', wsIds[row.y]);
                });
                let columnsByName = {};
                svColumns.forEach((svColumn) => { 
                    columnsByName[svColumn.name] = svColumn; 
                });    
                // have our own lib4x space 
                instance.options.lib4x = {
                    onDeleteRows : new Map(),
                    columnsByName: columnsByName,
                    dsMeta: dsMeta,
                    aggrEngine: hasAggregators(svAggregators) ? getAggrEngine(instance, svAggregators) : null
                };
                instance.options.lib4x.aggrEngine?.calculateAll();
                ctxPrototype = getCtxPrototype(instance); 
                let showHighlight = actionsContext.lookup('toggle-highlighting').showHighlight;
                // iterate the data to get the metadata for each row
                // and process the settings into markup classes (changes, highlight, issues)
                let svData = instance.getData();  
                svData.forEach((svRow, rowIndex) => {
                    let recordId = dsMetaUtil.getRecordId(instance, rowIndex);
                    let rowMeta = dsMeta.get(recordId);
                    if (rowMeta.gv.highlight)
                    {
                        setRowHighlightClass(instance, rowIndex, rowMeta, showHighlight);
                    }
                    // in case the model was not having data, a new row was inserted so also check ORIG_SV                    
                    if (rowMeta.gv.inserted || rowMeta.sv.inserted)
                    {
                        setRowInsertedClass(instance, rowIndex, rowMeta);
                    } 
                    else if (rowMeta.gv.updated)
                    {
                        setRowUpdatedClass(instance, rowIndex, rowMeta);
                    }  
                    if (rowMeta.issues)
                    {
                        setRowHasIssuesClass(instance, rowIndex, rowMeta);
                    }
                    let fields = rowMeta.gv.fields;
                    for (const fieldName in fields) {
                        if (fields[fieldName].changed || fields[fieldName].highlight || fields[fieldName].error || fields[fieldName].warning)
                        {
                            let colIndex = columnsByName[fieldName].index;
                            let cellName = jspreadsheet.helpers.getCellNameFromCoords(colIndex, rowIndex);
                            let cell = instance.getCell(cellName);
                            if (fields[fieldName].changed)
                            {
                                setCellChangedClass(cell, ORIG_GV, true);   
                            } 
                            if (fields[fieldName].highlight)
                            {
                                setCellHighlightClass(cell, fields[fieldName].highlight, true, showHighlight);   
                            }                             
                            if (fields[fieldName].error)
                            {
                                setCellErrorOrWarningClass(cell, fields[fieldName].message, 'error', true);   
                            }  
                            else if (fields[fieldName].warning)
                            {
                                setCellErrorOrWarningClass(cell, fields[fieldName].message, 'warning', true);   
                            }                           
                        }
                    }  
                });       
                // set any read only cells                                                   
                for (let cellName of wsReadOnlyCells)
                {
                    instance.setReadOnly(cellName, true);
                }
                $('#' + svStaticIdSv).data('jspreadsheet', instance.parent);
                // add hidden editors container
                $('#' + svStaticIdSv).append('<div class="' + C_LIB4X_SV_EDITORS_CONTAINER + ' u-vh" aria-hidden="true"></div>')
                // in case of dark theme, apply C_JSS_THEME_DARK class from jspreadsheet.themes.css
                if ($('body').hasClass(C_APEX_THEME_VITA_DARK)) {
                    $('#' + svStaticIdSv).addClass(C_JSS_THEME_DARK);
                }
                // prepare the toolbar, utilizing apex toolbar widget
                let toolbar$ = $('<div>').addClass('a-Toolbar ' + C_LIB4X_SV_TOOLBAR).insertAfter('#' + svStaticIdSv + ' .jtabs-headers-container');
                let action1controls = [];
                let action2controls = [];
                let action3controls = [];
                if (gridView.model.getOption('editable'))
                {
                    action1controls.push(
                        {
                            type: "BUTTON",
                            action: 'undo',
                            iconOnly: true
                        },   
                        {
                            type: "BUTTON",
                            action: 'redo',
                            iconOnly: true                         
                        }
                    );
                }
                if (gridView.model.allowAdd())
                {
                    action1controls.push(
                        {
                            type: "BUTTON",
                            action: 'add-row-before'
                        },
                        {
                            type: "BUTTON",
                            action: 'add-row-after'                           
                        }
                    );                    
                }
                let igConfig = igRegion.call('option').config;
                if (igConfig.editable && igConfig.editable.allowedOperations?.delete)
                {
                    action1controls.push(
                        {
                            type: "BUTTON",
                            action: 'delete-rows',
                        }
                    );      
                } 
                if (config.options.buttons.editOnFocus && gridView.model.getOption('editable'))
                {
                    if (!(igConfig.editable && igConfig.editable.allowedOperations?.create === false && igConfig.editable.allowedOperations?.update === false))
                    {
                        // by default, a double click is required to open the cell editor by mouse
                        // user can opt for and toggle edit on focus
                        // in case of edit on focus, shift key to be used to make a selection
                        action1controls.push(
                            {
                                type: 'TOGGLE',
                                action: 'toggle-edit-on-focus',
                                iconOnly: true
                            }
                        );   
                    }                 
                }                
                if (gridView.model.getOption('identityField'))    
                {   
                    // checking for identityField because when no primary key defined, model.fetchAll() will fail
                    // also in IG, when you scroll all the way down, it throws same error: 
                    // model.js?v=24.2.0:6425 Uncaught TypeError: Cannot read properties of undefined (reading 'length')
                    action2controls.push(
                        {
                            type: "BUTTON",
                            action: 'load-all',
                            iconBeforeLabel: true  
                        } 
                    );
                }
                action3controls.push(    
                    {
                        type: 'BUTTON',
                        action: 'show-help',
                        iconOnly: true                                
                    }
                );                 
                action3controls.push(
                    {
                        type: "BUTTON",
                        action: 'switch-pagination-type',
                        icon: worksheetOptions?.pagination ? 'fa fa-layout-nav-right-hamburger' : 'fa fa-button-group',
                        iconOnly: true 
                    } 
                );                 
                if (gridView.model.getOption('editable'))
                {                
                    action2controls.push(
                        {
                            type: "BUTTON",
                            action: 'synchronize',
                            iconOnly: true 
                        } 
                    );                
                    if (!worksheetOptions?.pagination)
                    {
                        action3controls.push(
                            {
                                type: 'RADIO_GROUP',
                                action: 'radiogroup-filter-rows',
                            }
                        );   
                    }                                                              
                }
                if (config.options.applyHighlighting)
                {
                    action3controls.push(    
                        {
                            type: 'TOGGLE',
                            action: 'toggle-highlighting',
                            iconOnly: true                                
                        }
                    );                    
                }                   
                let toolbarData = [
                    { id: 'actions1', controls: action1controls },
                    { id: 'actions2', controls: action2controls },
                    { id: 'actions3', controls: action3controls }
                ].filter(entry => entry.controls && entry.controls.length > 0);                
                toolbar$.toolbar({data: toolbarData, actionsContext: actionsContext});         
                toolbar$.find('[data-action="toggle-edit-on-focus"]').attr('data-no-update', true); // to prevent JS error
                toolbar$.find('[data-action="toggle-highlighting"]').attr('data-no-update', true); // to prevent JS error
                toolbar$.on('mousedown.svToolbar', function (e) {
                    e.stopImmediatePropagation();   // prevent the mousedown event to be picked up by JSS (would set instance.selectedRow to false)
                    // close any open editor before an action from toolbar is executed
                    // JSS would call closeEditor on blur, so alternative construct would be a finishEditing() construct like in IG.
                    // by closing the editor ourselves, the JSS blur handler gets nullified, so closing won't happen twice
                    closeAnyEditor(instance);
                }); 
                // have a toolbar resize observer as the toolbar height can change when toolbar items get wrapped
                let ro = new ResizeObserver(() => {
                    dlg$.trigger('dialogresize'); 
                });
                // take inner group as to prevent a ResizeObserver feedback loop
                let tbGroupContainer = toolbar$.find('.a-Toolbar-groupContainer');
                if (tbGroupContainer.length > 0)
                {
                    ro.observe(tbGroupContainer[0]);
                }
                // enable/disable actions              
                if (gridView.model.allowAdd())
                {
                    actionsContext.enable('add-row-before');
                    actionsContext.enable('add-row-after');
                }
                else
                {
                    actionsContext.disable('add-row-before');
                    actionsContext.disable('add-row-after');                    
                }  
                if (igConfig.editable && igConfig.editable.allowedOperations?.delete)
                {    
                    actionsContext.enable('delete-rows');    
                }
                else
                {
                    actionsContext.disable('delete-rows');    
                }  
                actionsContext.enable('load-all');
                actionsContext.enable('switch-pagination-type');
                hasHighlights ? actionsContext.show('toggle-highlighting') : actionsContext.hide('toggle-highlighting');
                try
                {
                    // move jss filter div with 'number of entries' select and/or search input into the toolbar as an extra toolbar group
                    let lastToolbarGroup$ = $('#' + svStaticIdSv + ' .a-Toolbar-groupContainer .a-Toolbar-group').last();
                    $('#' + svStaticIdSv + ' .jss_filter').insertAfter(lastToolbarGroup$);  
                    $('#' + svStaticIdSv + ' .jss_filter .jss_search').addClass('apex-item-text');
                    $('#' + svStaticIdSv + ' .jss_filter > div').addClass('a-Toolbar-item');
                    $('#' + svStaticIdSv + ' .jss_filter > div.a-Toolbar-item').filter(function () {
                        return $(this).children().length === 0;
                    }).remove();
                    $('#' + svStaticIdSv + ' .jss_filter').removeClass('jss_filter').addClass('a-Toolbar-group jss_filter_derived')
                    .find('.jss_pagination_dropdown').addClass('apex-item-select')
                    .parent().contents()
                    .filter(function () {
                        return this.nodeType === Node.TEXT_NODE;
                    }).remove();                
                    $('#' + svStaticIdSv + ' label:has(.jss_search)').contents().filter(function () {
                        return this.nodeType === Node.TEXT_NODE;
                    }).wrap('<span class="jss_search_label"></span>');   
                    $('#' + svStaticIdSv + ' .jss_search_label').text(getMessage('SEARCH'));    // Search 
                    $('#' + svStaticIdSv + ' .jss_search').after(
                        '<span class="a-Chip-clear js-clearInput jss_search_clear">' +
                            '<span class="a-Icon icon-multi-remove" aria-hidden="true"></span>' +
                        '</span>'
                    );
                    $('#' + svStaticIdSv + ' .jss_search_clear').hide();
                    $('#' + svStaticIdSv + ' .jss_search').on('input', function () {
                        let input$ = $(this);
                        let clear$ = $('#' + svStaticIdSv + ' .jss_search_clear');
                        input$.val().length > 0 ? clear$.show() : clear$.hide();
                    });
                    function resetSearch()
                    {
                        instance.resetSearch();
                        $('#' + svStaticIdSv + ' .jss_search_clear').hide();                                        
                    }
                    $('#' + svStaticIdSv + ' .jss_search').on('keydown', function (e) {
                        // prevent dialog close on escape; instead do a reset of the search
                        if (e.key === 'Escape') 
                        {
                            if ($(this).val().length > 0)
                            {
                                e.preventDefault();
                                resetSearch();
                            }
                        }
                    });                
                    $('#' + svStaticIdSv + ' .jss_search_clear').on('click', function(){
                        resetSearch();
                    });
                    $('#' + svStaticIdSv + ' .jss_pagination').children().first().css('visibility', 'hidden');
                }
                catch(e)
                {
                    $('.jss_filter').remove();
                    $('.jss_filter_derived').remove();
                }
                // dialog coordinates       
                // when the SV dialog was last time already opened on the same IG before, keep the
                // coordinates unchanged, else initialize to default settings
                if (setDialogCoordinates && ((igStaticId != lastClosedIgStaticId) || !config.options.rememberDialogCoordinates))
                {     
                    // make sure maxButton is reset as previous state could be maximized
                    let dlgWidget$ = dlg$.dialog("widget");  // will be the dialog wrapper
                    let maxButton$ = dlgWidget$.find('.ui-dialog-titlebar-max');
                    if (maxButton$ && maxButton$.hasClass(C_IS_MAX))
                    {
                        // Maximize
                        maxButton$.removeClass(C_IS_MAX).attr('title', getMessage('DIALOG.MAXIMIZE')); 
                    }                    
                    let heightFactor = 0.8;
                    dlg$.dialog('option', {
                        width: 'auto',
                        height: Math.floor(window.innerHeight * heightFactor)
                    });
                    // position left first as to make sure the centering will really work
                    dlg$.dialog('option', 'position', {
                        my: 'left',
                        at: 'left',
                        of: window
                    });
                    dlg$.dialog('option', 'position', {
                        my: 'center',
                        at: 'center',
                        of: window
                    });
                    dlg$.dialog('option', 'position', {
                        my: 'top',
                        at: 'top+'+Math.floor(window.innerHeight * (((1 - heightFactor) / 2) * 0.85)),
                        of: window
                    });                    
                }
                // restore any saved scrollbar positions
                let jssContent$ = dlg$.find('.jss_content');
                if (saveScrollTop)
                {
                    jssContent$.scrollTop(saveScrollTop);
                }  
                if (saveScrollLeft)
                {
                    jssContent$.scrollLeft(saveScrollLeft);
                }    
                // for page pagination, keep the same page in case of SV recreate
                if (worksheetOptions?.pagination && savePageNumber > 0)  
                {
                    instance.page(savePageNumber);
                }                       
                // commit edit on fill handle drag (else JSS will block the drag)
                $('#' + svStaticIdSv + ' .jss_corner').on('mousedown', function(event){
                    closeAnyEditor(instance);
                })       
                // set rows visibility as per current radiogroup setting
                let filterChoice = actionsContext.get('radiogroup-filter-rows');
                if (filterChoice != FILTER_ALL)
                {
                    setRowsVisibiliy(svStaticIdSv);
                }    
                // JSS sets tooltips to column headers equal to header text. Column tooltip prop is not working.
                // So we remove them like below
                $('#' + svStaticIdSv).find('.jss_worksheet thead td').removeAttr('title');                                                          
            }   // end createSpreadsheet function

            function hasAggregators(svAggregators) {
                if (!svAggregators || !Object.keys(svAggregators).length) {
                    return false;
                }
                return Object.values(svAggregators).some(a => a.length);
            }            

            function getFooterDefs(aggregators)
            {
                let footerDefs = [];
                // Total, Average, Minimum, Maximum
                let aggregates = [{type: 'SUM', label: getMessage('AGGREGATE.TOTAL')}, {type: 'AVG', label: getMessage('AGGREGATE.AVERAGE')}, {type: 'MIN', label: getMessage('AGGREGATE.MINIMUM')}, {type: 'MAX', label: getMessage('AGGREGATE.MAXIMUM')}];

                footerDefs = aggregates.filter(aggregate =>
                    Object.values(aggregators).some(colAggs =>
                        colAggs.some(a => a.type === aggregate.type)
                    )
                );    
                return footerDefs;     
            }

            function getAggrEngine(instance, aggregators)
            {
                let footerDefs = getFooterDefs(aggregators);  
                let footerRowIndex = {};
                footerDefs.forEach((fd, i) => {
                    footerRowIndex[fd.type] = i;
                });              

                function getAggregators()
                {
                    return aggregators;
                }

                function resetAggregators() 
                {
                    for (const colAggs of Object.values(aggregators)) 
                    {
                        for (const agg of colAggs) 
                        {
                            if ('sum' in agg)   agg.sum = 0;
                            if ('count' in agg) agg.count = 0;
                            if (agg.type === 'MIN') agg.value = Infinity;
                            else if (agg.type === 'MAX') agg.value = -Infinity;
                            else if ('value' in agg) agg.value = 0;
                        }
                    }
                }                

                function calculateAll() 
                {
                    resetAggregators();
                    for (const svRow of instance.getData()) 
                    {
                        for (const columnName in aggregators) 
                        {
                            let svColumn = instance.options.lib4x.columnsByName[columnName];
                            let rawValue = svRow[svColumn.index];
                            if (rawValue === '' || rawValue == null) continue;     // null check just added to be sure; JSS data is having empty values as '' though
                            let svValue = Number(rawValue);
                            if (isNaN(svValue)) continue;
                            aggregators[columnName].forEach(agg => {
                                if (agg.type === "SUM") 
                                {
                                    agg.value += svValue;
                                }
                                if (agg.type === "AVG") 
                                {
                                    agg.sum += svValue;
                                    agg.count++;
                                    agg.value = agg.sum / agg.count;
                                }
                                if (agg.type === "MIN") 
                                {
                                    agg.value = Math.min(agg.value, svValue);
                                }
                                if (agg.type === "MAX") 
                                {
                                    agg.value = Math.max(agg.value, svValue);
                                }
                            });
                        }
                    }
                    updateFooters();
                }   

                function applyChanges(changes, action) 
                {
                    let recalculateAll = false;
                    let updatedColumns = new Set();
                    changes.forEach(change => {
                        let svColumn = svColumns[change.x];
                        let colAggs = aggregators[svColumn.name];
                        if (!colAggs) return;
                        let newRaw = (action === 'undo') ? change.oldValue : change.value;
                        let oldRaw = (action === 'undo') ? change.value : change.oldValue;
                        let newIsValid = newRaw !== '' && newRaw != null && !isNaN(Number(newRaw));
                        let oldIsValid = oldRaw !== '' && oldRaw != null && !isNaN(Number(oldRaw));
                        let newVal = newIsValid ? Number(newRaw) : 0;
                        let oldVal = oldIsValid ? Number(oldRaw) : 0;
                        colAggs.forEach(agg => {
                            if (agg.type === "SUM") 
                            {
                                agg.value += newVal - oldVal;
                            }
                            if (agg.type === "AVG") 
                            {
                                // adjust sum
                                agg.sum += newVal - oldVal;
                                // adjust count
                                if (newIsValid && !oldIsValid) 
                                {
                                    agg.count++;
                                }
                                else if (!newIsValid && oldIsValid) 
                                {
                                    agg.count--;
                                }
                                agg.value = agg.count ? (agg.sum / agg.count) : 0;
                            }
                            if (agg.type === "MIN") 
                            {
                                if (newIsValid && newVal < agg.value) 
                                {
                                    agg.value = newVal;
                                }
                                else if (oldIsValid && oldVal === agg.value) 
                                {
                                    recalculateAll = true;
                                }
                            }
                            if (agg.type === "MAX") 
                            {
                                if (newIsValid && newVal > agg.value) 
                                {
                                    agg.value = newVal;
                                }
                                else if (oldIsValid && oldVal === agg.value) 
                                {
                                    recalculateAll = true;
                                }
                            }
                        });
                        updatedColumns.add(svColumn.name);
                    });
                    if (recalculateAll)
                    {
                        calculateAll();
                    }
                    else
                    {
                        updateFooters(updatedColumns);
                    }
                }                
                
                function updateFooters(updatedColumns)
                {
                    for (const columnName in aggregators) 
                    {
                        if (!updatedColumns || updatedColumns.has(columnName))
                        {
                            let svColumn = instance.options.lib4x.columnsByName[columnName];
                            aggregators[columnName].forEach(agg => {
                                let rowIndex = footerRowIndex[agg.type];
                                let formattedAggrValue = apex.locale.formatNumber(agg.value, svColumn.customFormat).trim();  
                                instance.options.footers[rowIndex][svColumn.index] = formattedAggrValue;
                                $(instance.tfoot).find('tr').eq(rowIndex).find('td').eq(svColumn.index + 1).text(formattedAggrValue);  // +1 for row number column
                            });
                        }
                    }                         
                }

                return {
                    getAggregators: getAggregators,
                    calculateAll: calculateAll,
                    applyChanges: applyChanges
                }
            }
            
            // parse dateStringValue, taking into account any configured alternative date masks
            function dateInputToDataValue(targetFormatMask, dateStringValue)
            {
                let altMasks = config.options.alternativeDateMasks ?? {};

                let result = dateStringValue;
                let masksArray = altMasks.hasOwnProperty(targetFormatMask) ? [...altMasks[targetFormatMask]] : [];
                function pushAlternative(find, replace)
                {
                    if (targetFormatMask.includes(find))
                    {
                        masksArray.push(targetFormatMask.replace(find, replace));
                    }
                }
                pushAlternative('MON', 'MM');
                pushAlternative('MM', 'MON');
                pushAlternative('YYYY', 'RR');
                pushAlternative('RR', 'YYYY');

                let inclTime = targetFormatMask.includes('HH');
                // add generic format to the start of te array
                masksArray.unshift('YYYY-MM-DD' + (inclTime ? ' HH24:MI:SS' : ''));

                // add targetFormatMask to the start of the array 
                masksArray.unshift(targetFormatMask);
                if (dateStringValue)
                {
                    let parsedDate = null;
                    let dateOk = false;
                    let masksIndex = 0;
                    let dateMask = masksArray[masksIndex];
                    // iterate all masks until any successfull parse
                    while ((masksIndex < masksArray.length) && !dateOk)
                    {
                        try 
                        {
                            if (dateMask == 'DD')
                            {
                                parsedDate = new Date();
                                let currentMonth = parsedDate.getMonth();
                                let currentYear = parsedDate.getFullYear();
                                let setResult = parsedDate.setDate(dateStringValue);   // set day
                                if (isNaN(setResult) || (parsedDate.getMonth() != currentMonth) || (parsedDate.getFullYear() != currentYear))
                                {
                                    throw new Error('Entered value is not a valid day');
                                }
                                dateOk = true;
                            } 
                            else if ((dateMask == 'DD/MM') || (dateMask == 'MM/DD') || (dateMask == 'DD-MM') || (dateMask == 'MM-DD') || (dateMask == 'DD.MM') || (dateMask == 'MM.DD'))
                            {
                                let currentYear = new Date().getFullYear();
                                // whatever the separator, the below with adding using '/' just works 
                                parsedDate = apex.date.parse(dateStringValue + '/' + currentYear, dateMask + '/YYYY');
                                dateOk = true;
                            }  
                            else
                            {                         
                                parsedDate = apex.date.parse(dateStringValue, dateMask);
                                dateOk = true;
                            }
                        }
                        catch (e) 
                        {
                            dateOk = false;
                            masksIndex++;
                            dateMask = masksArray[masksIndex];
                        }
                    }
                    if (dateOk && parsedDate)
                    {
                        // successfully parsed - now format to the target format mask
                        // intentionally also when maskIndex == 0 as a date like 
                        // 0811-2023 gets successfully parsed with mask 'DS' but then the
                        // formattedValue will be 8/01/2023
                        result = apex.date.format(parsedDate, targetFormatMask);
                    }
                }
                return result;           
            }            

            // get any effectively configured msoNumberFormat for the column
            // where the config can be a value or a function
            function getConfiguredMsoNumberFormat(svColumn, formatConfig)
            {
                let msoNumberFormat = null;
                if (svColumn.lib4x?.msoNumberFormat)
                {
                    // resolve from value or function
                    msoNumberFormat = util.resolveValue(svColumn.lib4x.msoNumberFormat, svColumn.customFormat);
                }
                else if (formatConfig)
                {
                    msoNumberFormat = util.resolveValue(formatConfig, svColumn.name, svColumn.customFormat);
                } 
                return msoNumberFormat;
            }

            // block with editor interfaces
            let svEditors = function() {
                let numberInterface = function() {
                    let methods = {};      
                    methods.createCell = function (cell, value, x, y, instance, options) {
                        $(cell).text(value);
                        return value;                        
                    }  

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        let formattedValue = options.lib4x.toFormattedValue(value);
                        $(cell).text(formattedValue);
                        return value;                        
                    }        
                    
                    methods.openEditor = function (cell, value, x, y, instance, options, ctx) {
                        let clientRectInfo = cell.getBoundingClientRect();                        
                        let cell$ = $(cell);
                        cell$.addClass(C_EDITOR); // JSS editor class
                        cell$.addClass(C_SV_EDITOR);
                        let editor$ = options.lib4x.editor$;
                        let apexItem = options.lib4x.apexItem;
                        let input$ = null;
                        if (!editor$) {
                            editor$ = $('<div class="' + C_LIB4X_SV_CELL_EDITOR + ' a-GV-columnItem"></div>'); 
                            input$ = apexItem.element.clone();
                            input$.attr('id', input$.attr('id') + '_sv').attr('name', input$.attr('name') + '_sv');
                            editor$.append(input$);
                            options.lib4x.editor$ = editor$;
                        }
                        else {
                            input$ = editor$.find('input');
                        }
                        cell$.empty().append(editor$);
                        let formattedValue = options.lib4x.toFormattedValue(value);
                        input$.val(formattedValue);
                        let isClosing = false;
                        input$.off('keydown.lib4x').on('keydown.lib4x', function (e) {
                            if (e.key === 'Tab' || e.key === 'Enter' || e.key === 'Escape') { 
                                e.preventDefault();
                                // it happens rarely, but it can happen 2 key events do come in
                                //if (isClosing) return;  // extra safeguard - looks like not needed though
                                isClosing = true;                                
                                if (instance.edition != null)
                                {    
                                    if ($(cell).hasClass(C_EDITOR))
                                    {                          
                                        instance.closeEditor(cell, e.key !== 'Escape');
                                    }
                                }                                
                            }
                        });
                        input$.css({
                            width: (clientRectInfo.width + 2) + 'px',
                            height: (clientRectInfo.height - 2) + 'px'
                        });
                        if (util.isPrintableKey(ctx)) 
                        {
                            input$.val(ctx.key == ' ' ? '' : ctx.key);
                        }                     
                        setTimeout(() => {
                            input$.focus();
                            if (ctx && ctx instanceof MouseEvent)
                            {
                                let len = input$.val().length;
                                input$[0].setSelectionRange(len, len);
                            }
                        }, 0);              
                    }     
                    
                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        //let apexItem = options.lib4x.apexItem;
                        // the C_EDITOR class is removed by JSS already
                        $(cell).removeClass(C_SV_EDITOR);
                        let input$ = $(cell).find('input');
                        //input$.off('blur.lib4x')
                        input$.off('keydown.lib4x')
                        let value = input$.val();
                        let editor$ = options.lib4x.editor$;
                        $('#' + svStaticIdSv + ' .' + C_LIB4X_SV_EDITORS_CONTAINER).append(editor$);
                        if (confirmChanges) 
                        {        
                            let dataValue = options.lib4x.inputToDataValue(value);     
                            let formattedValue = options.lib4x.toFormattedValue(dataValue);           
                            cell.textContent = formattedValue;
                            let currentDataValue = instance.getValueFromCoords(x, y, false);
                            if (currentDataValue === dataValue)
                            {
                                return currentDataValue;    // returning the original value to prevent a setValue (plus events) to get executed by JSS
                            }
                            return dataValue;
                        }
                        // in case confirmChanges is false, JSS will discard any return value and reset the cell content to the original content (instance.edition[1])
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    // below method is mentioned in the documentation, but not sure when it is called if at all
                    // it is not called when copying to clipboard or when using 'Save As..'
                    // also couldn't find any call in the github repository (type.get)
                    methods.get = function (options, value) {
                        return value;
                    }                    

                    return methods;
                }();                    

                let selectListInterface = function() {
                    let methods = {};

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        $(cell).text(value?.d);
                        $(cell).addClass('a-GV-cell');
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        if (value && (typeof value === 'object') && value.hasOwnProperty('d')) {
                            $(cell).text(value.d);
                        }
                        else {
                            $(cell).empty();
                            value = { v: '', d: '' };
                        }
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options) {
                        let clientRectInfo = cell.getBoundingClientRect();                        
                        let cell$ = $(cell);
                        cell$.addClass(C_EDITOR); // JSS editor class
                        cell$.addClass(C_SV_EDITOR);
                        let editor$ = options.lib4x.editor$;
                        let apexItem = options.lib4x.apexItem;
                        let itemType = apexItem.item_type;
                        let select$ = null;
                        if (!editor$) {
                            editor$ = $('<div class="' + C_LIB4X_SV_CELL_EDITOR + ' a-GV-columnItem"></div>');     // by intention a-GV-columnItem added for proper styling
                            if (itemType == 'SELECT')
                            {
                                select$ = apexItem.element.clone();
                            }
                            else if (itemType == 'RADIO_GROUP')
                            {
                                select$ = deriveSelectFromRadioGroup(apexItem);
                            }
                            select$.attr('id', select$.attr('id') + '_sv').attr('name', select$.attr('name') + '_sv');
                            editor$.append(select$);
                            options.lib4x.editor$ = editor$;
                        }
                        else {
                            select$ = editor$.find('select');
                        }
                        cell$.empty().append(editor$);
                        select$.val(value?.v);
                        let isClosing = false;
                        select$.off('keydown.svSelectNav').on('keydown.svSelectNav', function (e) {
                            // prevent close dialog (escape), etc behavior 
                            // instead have the JSS behavior effectively and close the editor
                            if (e.key === 'Tab' || e.key === 'Enter' || e.key === 'Escape') {  
                                e.preventDefault();
                                //if (isClosing) return;            // extra safety guard, but looks like not needed
                                isClosing = true;                                  
                                if (instance.edition != null)
                                {
                                    if ($(cell).hasClass(C_EDITOR))
                                    {                                       
                                        instance.closeEditor(cell, e.key !== 'Escape');
                                    }
                                }
                            }
                        });
                        select$.css({
                            width: (clientRectInfo.width + 2) + 'px',
                            height: (clientRectInfo.height - 2) + 'px'
                        });
                        setTimeout(() => select$.focus(), 0);
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        let apexItem = options.lib4x.apexItem;
                        // the C_EDITOR class is removed by JSS already
                        $(cell).removeClass(C_SV_EDITOR);
                        let value = $(cell).find('select').val();
                        let editor$ = options.lib4x.editor$;
                        $('#' + svStaticIdSv + ' .' + C_LIB4X_SV_EDITORS_CONTAINER).append(editor$);
                        if (confirmChanges) 
                        {                        
                            let displayValue = apexItem.displayValueFor(value);
                            cell.textContent = displayValue;
                            let currentDataValue = instance.getValueFromCoords(x, y, false);
                            if (currentDataValue?.v === value)
                            {
                                return currentDataValue;    // returning the original object to prevent a setValue (plus events) to get executed by JSS
                            }
                            return {v: value, d: displayValue};
                        }
                        // in case confirmChanges is false, JSS will discard any return value and reset the cell content to the original content (instance.edition[1])
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    // below method is mentioned in the documentation, but not sure when it is called if at all
                    // it is not called when copying to clipboard or when using 'Save As..'
                    // also couldn't find any call in the github repository (type.get)
                    methods.get = function (options, value) {
                        return value;
                    }

                    return methods;
                }();

                let dateInterface = function () {
                    let methods = {};

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        $(cell).text(value);
                        $(cell).addClass('a-GV-cell');
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        $(cell).text(value);
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options, ctx) {
                        let apexItem = options.lib4x.apexItem;
                        let clientRectinfo = cell.getBoundingClientRect();                        
                        let cell$ = $(cell);
                        cell$.addClass(C_EDITOR);       // JSS editor class
                        cell$.addClass(C_SV_EDITOR);    // used in mousedown global event handler
                        let editor$ = options.lib4x.editor$;
                        let datePicker$ = null;
                        if (!editor$) {
                            editor$ = $('<div class="' + C_LIB4X_SV_CELL_EDITOR + ' a-GV-columnItem"></div>');     // by intention a-GV-columnItem added for proper styling
                            datePicker$ = util.item.cloneWebComponent(apexItem.id, '_sv');
                            editor$.append(datePicker$);
                            options.lib4x.editor$ = editor$;
                        }
                        else {
                            datePicker$ = editor$.find('a-date-picker');
                        }
                        cell$.empty().append(editor$);
                        // when the editor is newly created, only after append, the input element will be there 
                        let input$ = datePicker$.find('input');
                        // (re)establish the keydown eventhandler as per current cell, x, y values 
                        let isClosing = false;                             
                        input$.off('keydown.svDatePickerNav').on('keydown.svDatePickerNav', function (e) {
                            if (e.key === 'Tab' || e.key === 'Enter' || e.key === 'Escape') { 
                                e.preventDefault();
                                // it happens rarely, but it can happen 2 key events do come in
                                //if (isClosing) return;  // extra safeguard - looks like not needed though
                                isClosing = true;                                
                                if (instance.edition != null)
                                {    
                                    if ($(cell).hasClass(C_EDITOR))
                                    {                          
                                        instance.closeEditor(cell, e.key !== 'Escape');
                                    }
                                }                                
                            }
                        });      

                        datePicker$[0].setValue(value);
                        let buttonWidth = datePicker$.find('.a-Button--calendar').outerWidth();
                        datePicker$.css({
                            width: (clientRectinfo.width + 2) + 'px',
                            height: (clientRectinfo.height - 2) + 'px'
                        });
                        input$.css({
                            width: (clientRectinfo.width - buttonWidth - 1) + 'px',
                            height: (clientRectinfo.height - 2) + 'px',
                            minHeight: (clientRectinfo.height - 2) + 'px'
                        });
                        if (util.isPrintableKey(ctx)) 
                        {
                            input$.val(ctx.key == ' ' ? '' : ctx.key);
                        }                     
                        setTimeout(() => {
                            input$.focus();
                            if (ctx && ctx instanceof MouseEvent)
                            {
                                let len = input$.val().length;
                                input$[0].setSelectionRange(len, len);
                            }
                        }, 0);
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        let value = $(cell).find('input').val();
                        let editor$ = options.lib4x.editor$;
                        $('#' + svStaticIdSv + ' .' + C_LIB4X_SV_EDITORS_CONTAINER).append(editor$);
                        // the C_EDITOR class is removed by JSS already
                        $(cell).removeClass(C_SV_EDITOR);
                        $('.ui-dialog-datepicker:visible').hide();                        
                        if (confirmChanges) {
                            cell.textContent = value;
                            return value;
                        }
                    }

                    methods.destroyCell = function (cell, x, y, instance, options) {

                    }

                    methods.get = function (options, value) {
                        return value;
                    }

                    return methods;
                }(); 

                let checkboxInterface = function () {
                    let methods = {};

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        let checkedAttr = (value == options.lib4x.checkedValue) ? 'checked' : '';
                        let cbHtml$ = $('<div class="apex-item-single-checkbox"><input type="checkbox" ' + checkedAttr + ' tabindex="0"><span class="u-checkbox" aria-hidden="true"></span><span class="u-vh" aria-hidden="true"></span></div>');
                        $(cell).empty().append(cbHtml$);
                        $(cell).addClass('a-GV-cell');  // using this class, the checkbox reacts better to clicking
                        $(cell).on('click', 'input', function (jQueryEvent) {
                            jQueryEvent.preventDefault();
                            jQueryEvent.stopPropagation();
                            if ((options?.readOnly === true) || (instance.isReadOnly(x, y))) {
                                return false;
                            }
                            let currentValue = this.checked ? options.lib4x.checkedValue : options.lib4x.uncheckedValue;
                            setTimeout(function () {
                                instance.setValue(cell, currentValue);
                            });
                        });
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        // to be sure, check for false/true value, shouldn't happen though
                        if (value === false || value === true) return;
                        if (cell) {
                            $(cell).find('input[type="checkbox"]').prop('checked', (value == options.lib4x.checkedValue));
                            $(cell).find('.u-vh').text(options.lib4x.label[value]);
                        }
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options) {
                        instance.closeEditor(cell, false);
                        value = (value == options.lib4x.checkedValue) ? options.lib4x.uncheckedValue : options.lib4x.checkedValue;
                        setTimeout(function () {
                            instance.setValue(cell, value);
                        });
                        return false;
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        let isChecked = $(cell).find('input[type="checkbox"]').prop('checked');
                        let value = isChecked ? options.lib4x.checkedValue : options.lib4x.uncheckedValue;
                        return value;
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    methods.get = function (options, value) {
                        return options.lib4x.label[value];
                    }

                    return methods;
                }();

                let switchInterface = function () {
                    let methods = {};

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        let apexItem = options.lib4x.apexItem;
                        let cbHtml$ = apexItem.element.closest('.a-Switch').clone();
                        let input$ = cbHtml$.find('input');
                        let checked = value?.v == options.lib4x.onValue;
                        input$.removeAttr('id name checked aria-describedby').attr('tabindex', '-1');
                        input$.prop('checked', checked);
                        $(cell).empty().append(cbHtml$);
                        $(cell).addClass('a-GV-cell');
                        $(cell).on('click', 'input, span.a-Switch', function (jQueryEvent) {
                            jQueryEvent.preventDefault();
                            jQueryEvent.stopPropagation();
                            if ((options?.readOnly === true) || (instance.isReadOnly(x, y))) {
                                return false;
                            }
                            let cbInput$ = $(cell).find('input');
                            if (!$(jQueryEvent.target).is(':checkbox'))
                            {
                                cbInput$.prop('checked', !cbInput$.prop('checked'));
                            }
                            let currentValue = cbInput$[0].checked ? options.lib4x.onValue : options.lib4x.offValue;
                            currentValue = {v: currentValue, d: options.lib4x.label[currentValue]};
                            setTimeout(function () {
                                instance.setValue(cell, currentValue);
                            });
                        });
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        // to be sure, check for false/true value, shouldn't happen though
                        if (value === false || value === true) return;
                        if (cell) {
                            $(cell).find('input[type="checkbox"]').prop('checked', (value?.v == options.lib4x.onValue));
                        }
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options) {
                        instance.closeEditor(cell, false);
                        let newValue = (value?.v == options.lib4x.onValue) ? options.lib4x.offValue : options.lib4x.onValue;
                        newValue = {v: newValue, d: options.lib4x.label[newValue]};
                        setTimeout(function () {
                            instance.setValue(cell, newValue);
                        });
                        return false;
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        let isChecked = $(cell).find('input[type="checkbox"]').prop('checked');
                        let value = isChecked ? options.lib4x.onValue : options.lib4x.offValue;
                        return value;
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    methods.get = function (options, value) {
                        return (value == options.lib4x.onValue) ? options.lib4x.onValue : options.lib4x.offValue;
                    }

                    return methods;
                }();  

                let simpleRadioInterface = function () {
                    let methods = {};
                    let rgSeq = 0;

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        let apexItem = options.lib4x.apexItem;
                        let rgHtml$ = apexItem.element.clone();
                        rgSeq = rgSeq + 1;
                        rgHtml$.attr('id', rgHtml$.attr('id') + '_sv' + rgSeq);
                        rgHtml$.addClass('apex-item-single-checkbox');
                        if (rgHtml$.find('.apex-item-grid').length == 0)
                        {
                            // item grid missing, probably because the column was configured with a number of columns being 1
                            // let's insert it as we want horizontal orientation only
                            rgHtml$.removeAttr('aria-orientation');
                            let grid$ = $('<div>', {class: 'apex-item-grid radio_group', role: 'none'});
                            var gridRow$ = $('<div>', {class: 'apex-item-grid-row', role: 'none'});
                            rgHtml$.children('.apex-item-option').appendTo(gridRow$);
                            grid$.append(gridRow$);
                            rgHtml$.append(grid$);                                    
                        }
                        function adjustRadioOption(seqNo, optionValue)
                        {
                            let radioOption$ = rgHtml$.find('.apex-item-option').eq(seqNo).find('input[type="radio"]');
                            let optionId = radioOption$.attr('id') + '_sv' + rgSeq;
                            let optionName = radioOption$.attr('name') + '_sv' + rgSeq;
                            radioOption$.attr('id', optionId).attr('name', optionName);
                            radioOption$.parent().find('label').attr('for', optionId);
                            radioOption$[0].checked = (value?.v == optionValue);                                    
                        }
                        adjustRadioOption(0, options.lib4x.firstValue);
                        adjustRadioOption(1, options.lib4x.secondValue);
                        $(cell).empty().append(rgHtml$);
                        $(cell).addClass('a-GV-cell'); 
                        $(cell).on('click', 'input', function (jQueryEvent) {
                            if ((options?.readOnly === true) || (instance.isReadOnly(x, y))) {
                                jQueryEvent.preventDefault();
                                jQueryEvent.stopPropagation();                                        
                                return false;
                            }
                        });                                
                        $(cell).on('change', 'input[type=radio]', (e) => {   
                            if ((options?.readOnly === true) || (instance.isReadOnly(x, y))) {
                                return false;
                            }   
                            let radioButton$ = $(e.target);
                            if (radioButton$[0].checked)
                            {
                                setTimeout(function () {
                                    let currentValue = radioButton$.val();
                                    currentValue = {v: currentValue, d: options.lib4x.label[currentValue]};
                                    instance.setValue(cell, currentValue);
                                });                                        
                            }
                        });                                
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        // to be sure, check for false/true value, shouldn't happen though
                        if (value === false || value === true) return;
                        if (cell) {
                            function setOptionState(seqNo, optionValue)
                            {
                                $(cell).find('.apex-item-option').eq(seqNo).find('input[type="radio"]')[0].checked = (value?.v == optionValue);
                            }
                            setOptionState(0, options.lib4x.firstValue);
                            setOptionState(1, options.lib4x.secondValue);
                        }
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options) {
                        instance.closeEditor(cell, false);
                        let newValue = (value?.v == options.lib4x.firstValue) ? options.lib4x.secondValue : options.lib4x.firstValue;
                        newValue = {v: newValue, d: options.lib4x.label[newValue]};
                        setTimeout(function () {
                            instance.setValue(cell, newValue);
                        });
                        return false;
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        let value = $(cell).find('input[type=radio]:checked').val() || null;
                        return value ? {v: value, d: options.lib4x.label[value]} : {v: '', d: ''};
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    methods.get = function (options, value) {
                        return value;
                    }

                    return methods;
                }();

                let pillButtonsInterface = function () {
                    let methods = {};
                    let rgSeq = 0;

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        let apexItem = options.lib4x.apexItem;
                        let rgHtml$ = apexItem.element.clone();
                        rgSeq = rgSeq + 1;
                        rgHtml$.attr('id', rgHtml$.attr('id') + '_sv' + rgSeq);
                        rgHtml$.addClass('apex-item-grid-row');
                        function adjustRadioOption(seqNo, optionValue)
                        {
                            let radioOption$ = rgHtml$.find('.apex-item-option').eq(seqNo).find('input[type="radio"]');
                            let optionId = radioOption$.attr('id') + '_sv' + rgSeq;
                            let optionName = radioOption$.attr('name') + '_sv' + rgSeq;
                            radioOption$.attr('id', optionId).attr('name', optionName);
                            radioOption$.parent().find('label').attr('for', optionId);
                            radioOption$[0].checked = (value?.v == optionValue);                                    
                        }
                        adjustRadioOption(0, options.lib4x.firstValue);
                        adjustRadioOption(1, options.lib4x.secondValue);
                        $(cell).empty().append(rgHtml$);
                        $(cell).addClass('a-GV-cell'); 
                        $(cell).on('click', 'input', function (jQueryEvent) {
                            if ((options?.readOnly === true) || (instance.isReadOnly(x, y))) {
                                jQueryEvent.preventDefault();
                                jQueryEvent.stopPropagation();                                        
                                return false;
                            }
                        });                                 
                        $(cell).on('change', 'input[type=radio]', (e) => {                                     
                            if ((options?.readOnly === true) || (instance.isReadOnly(x, y))) {
                                return false;
                            }  
                            let radioButton$ = $(e.target);
                            if (radioButton$[0].checked)
                            {
                                setTimeout(function () {
                                    let currentValue = radioButton$.val();
                                    currentValue = {v: currentValue, d: options.lib4x.label[currentValue]};
                                    instance.setValue(cell, currentValue);
                                });                                        
                            }
                        });
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        // to be sure, check for false/true value, shouldn't happen though
                        if (value === false || value === true) return;
                        if (cell) {
                            function setOptionState(seqNo, optionValue)
                            {
                                $(cell).find('.apex-item-option').eq(seqNo).find('input[type="radio"]')[0].checked = (value?.v == optionValue);
                            }
                            setOptionState(0, options.lib4x.firstValue);
                            setOptionState(1, options.lib4x.secondValue);
                        }
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options) {
                        instance.closeEditor(cell, false);
                        let newValue = (value?.v == options.lib4x.firstValue) ? options.lib4x.secondValue : options.lib4x.firstValue;
                        newValue = {v: newValue, d: options.lib4x.label[newValue]};
                        setTimeout(function () {
                            instance.setValue(cell, newValue);
                        });
                        return false;
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        let value = $(cell).find('input[type=radio]:checked').val() || null;
                        return value ? {v: value, d: options.lib4x.label[value]} : {v: '', d: ''};
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    methods.get = function (options, value) {
                        return value;
                    }

                    return methods;
                }(); 

                let selectOneInterface = function() {
                    let methods = {};

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        $(cell).text(value?.d);
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        if (value && (typeof value === 'object') && value.hasOwnProperty('d')) {
                            $(cell).text(value.d);
                        }
                        else {
                            $(cell).empty();
                            value = { v: '', d: '' };
                        }
                        return value;
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options, ctx) {
                        let apexItem = options.lib4x.apexItem;
                        let clientRectInfo = cell.getBoundingClientRect();
                        let cell$ = $(cell);
                        cell$.addClass(C_EDITOR); // JSS editor class
                        cell$.addClass(C_SV_EDITOR);  // used in mousedown global event handler
                        let editor$ = options.lib4x.editor$;
                        let select$ = null;
                        if (!editor$) {
                            editor$ = $('<div class="' + C_LIB4X_SV_CELL_EDITOR + '"></div>');
                            let aOptions = Array.from(options.lib4x.lovMap)
                                .sort((a, b) => a[0].localeCompare(b[0])) // Sort alphabetically by display name (d)
                                .map(([d, r]) => `<a-option value="${r}">${d}</a-option>`)
                                .join('');                            
                            // max-results attribute is supported; not used
                            select$ = $('<a-select id="' + apexItem.id + '_sv" return-display="true" min-characters-search="0" match-type="contains">' + aOptions + '</a-select>');
                            editor$.append(select$);
                            options.lib4x.editor$ = editor$;
                        }
                        else {
                            select$ = editor$.find('a-select');
                        }
                        cell$.empty().append(editor$);
                        select$[0].setValue(value?.v);
                        let isClosing = false;   
                        select$.off('keydown.svSelectNav').on('keydown.svSelectNav', function (e) {
                            if (e.key === 'Tab' || e.key === 'Enter' || e.key === 'Escape') { 
                                e.preventDefault(); 
                                // if (isClosing) return;           // safety guard
                                isClosing = true;                                  
                                if (instance.edition != null)
                                {
                                    if ($(cell).hasClass(C_EDITOR))
                                    {                                      
                                        instance.closeEditor(cell, e.key !== 'Escape');
                                    }
                                }                                                              
                            }                         
                        });
                        select$.find('.apex-item-comboselect').css({
                            height: (clientRectInfo.height - 2) + 'px'
                        });
                        setTimeout(() => {
                            let input$ = select$.find('input');
                            input$.focus();
                            if (ctx && ctx instanceof MouseEvent)
                            {
                                let len = input$.val().length;
                                input$[0].setSelectionRange(len, len);
                            }
                        }, 0);                        
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {
                        // the C_EDITOR class is removed by JSS already
                        $(cell).removeClass(C_SV_EDITOR);
                        let displayValue = util.normalizeDisplayValue($(cell).find('a-select').find('input').val());
                        let value = options.lib4x.lovMap.get(displayValue) || '';
                        if (!value)
                        {
                            displayValue = '';
                        }
                        let editor$ = options.lib4x.editor$;
                        $('#' + svStaticIdSv + ' .' + C_LIB4X_SV_EDITORS_CONTAINER).append(editor$);
                        $('.ui-dialog-combobox:visible').hide();                        
                        if (confirmChanges) {
                            cell.textContent = displayValue;
                            let currentDataValue = instance.getValueFromCoords(x, y, false);
                            if (currentDataValue?.v === value)
                            {
                                return currentDataValue;    // returning the original object to prevent a setValue (plus events) to get executed by JSS
                            }
                            return {v: value, d: displayValue};
                        }
                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    methods.get = function (options, value) {
                        return value;
                    }

                    return methods;
                }();      
                
                // for any future use
                let noeditInterface = function () {
                    let methods = {};

                    methods.createCell = function (cell, value, x, y, instance, options) {
                        let displayValue = util.getDisplayValue(value) || '';
                        $(cell).empty().text(displayValue);
                        $(cell).addClass('lib4x-SV-noedit');
                        return value;
                    }

                    methods.updateCell = function (cell, value, x, y, instance, options) {
                        if (value && (typeof value === 'object') && value.hasOwnProperty('d')) {
                            $(cell).text(value.d);
                        }
                        else {
                            $(cell).empty();
                            value = { v: '', d: '' };
                        }
                        return value;                                
                    }

                    methods.openEditor = function (cell, value, x, y, instance, options) {
                        instance.closeEditor(cell, false);
                        return false;
                    }

                    methods.closeEditor = function (cell, confirmChanges, x, y, instance, options) {

                    }

                    methods.destroyCell = function (cell, x, y, instance) {

                    }

                    methods.get = function (options, value) {
                        return value;
                    }

                    return methods;
                }();                

                return {
                    numberInterface: numberInterface,
                    selectListInterface: selectListInterface,
                    dateInterface: dateInterface,
                    checkboxInterface: checkboxInterface,
                    switchInterface: switchInterface,
                    simpleRadioInterface: simpleRadioInterface,
                    pillButtonsInterface: pillButtonsInterface,
                    selectOneInterface: selectOneInterface,
                    noeditInterface: noeditInterface                       
                }
            }();

            // compose the columns for the spreadsheetView as basically derived from IG columns                                 
            let svColumns = [];
            let colIndex = 0;
            let igColumns = gridView.view$.grid('getColumns');
            let igColumnGroups = gridView.view$.grid('option', 'columnGroups');
            let userAggregates = apex.region(igStaticId).call('getAggregates')?.filter((agg) => agg.columnType=='column' && agg.isEnabled);
            let svAggregators  = {};
            for (columnNo in igColumns) {
                let igColumn = igColumns[columnNo]; 
                let apexItem = apex.item(igColumn.elementId);
                let itemType = apexItem.item_type;
                // skip condition
                let skip = ((!igColumn.elementId) ||
                    (igColumn.dependsOn && igColumn.calcValue && config.options.excludeCalculatedColumns) || 
                    (igColumn.hidden) ||
                    (igColumn.cellTemplate) ||   // HTML Expressions not supported
                    (apexItem.element.attr('type') == 'password') ||
                    (apexItem.item_type == 'PCT_GRAPH') ||
                    //(igColumn.linkAttributes && igColumn.linkAttributes.includes(C_LIB4X_SV_EXCLUDE)) ||    // was having this before for link columns
                    (apexItem.item_type == 'IG_BUTTON_COLUMN') ||   
                    (apexItem.item_type == 'IG_SIMPLE_CHOICE') ||   
                    (igColumn.lib4x?.spreadsheetViewOptions?.exclude) ||  
                    // the C_LIB4X_SV_EXCLUDE class can be given on the column (appearance section) as to indicate the column should not be included in the SV                    
                    (igColumn.columnCssClasses && igColumn.columnCssClasses.split(/\s+/).includes(C_LIB4X_SV_EXCLUDE)));
                if (!skip) {
                    let svColumn = {};
                    svColumn.index = colIndex;
                    colIndex = colIndex + 1;
                    svColumn.lib4x = {};
                    svColumn.name = igColumn.property;
                    svColumn.title = igColumn.label;    // is either 'Alternative Label' or 'Heading'
                    svColumn.type = 'text'; // default to text
                    svColumn.width = igColumn.curWidth * 0.97;
                    svColumn.align = { start: 'left', end: 'right' }[igColumn.alignment] ?? 'center';
                    svColumn.lib4x.apexItem = apexItem;
                    svColumn.lib4x.defaultValue = igColumn.defaultValue;
                    if (igColumn.groupName && igColumn.useGroupFor?.includes('heading'))
                    {
                        svColumn.lib4x.columnGroup = igColumn.groupName;
                    }
                    // default readonly
                    svColumn.readOnly = true;
                    if (itemType == 'NUMBER' && (igColumn.aggregates || userAggregates?.length > 0))
                    {
                        // if user has defined any aggregates, use these, else use any aggregates as defined on column properties init function
                        if (userAggregates?.length > 0)
                        {
                            svColumn.lib4x.aggregates = userAggregates.filter(aggr => aggr.columnId == igColumn.id).map(aggr => aggr.function);
                        }
                        else
                        {
                            // make copy but skip custom aggregates (are defined by objects)
                            svColumn.lib4x.aggregates = igColumn.aggregates.filter(aggr => typeof aggr === 'string');
                        }
                        if (svColumn.lib4x.aggregates.length > 0)
                        {
                            // Aggregator Model
                            svAggregators[svColumn.name] = [
                                { type: "SUM", value: 0 },
                                { type: "AVG", sum: 0, count: 0, value: 0 },
                                { type: "MIN", value: Infinity },
                                { type: "MAX", value: -Infinity }
                            ].filter((aggregator) => svColumn.lib4x.aggregates.includes(aggregator.type));
                        }
                    }
                    svColumn.render = function (td, value, x, y, worksheet, options) {
                        if (td && value) {  
                            // value might be object with return and display values                      
                            td.textContent = util.getDisplayValue(value);
                            // if there are ever complaints about html content not rendering,
                            // igColumn.escape property can be taken into consideration
                        }
                    }                            
                    if (igColumn.lib4x?.spreadsheetViewOptions?.updateGridColumnAfter)
                    {
                        svColumn.lib4x.updateGridColumnAfter = util.toArray(igColumn.lib4x.spreadsheetViewOptions.updateGridColumnAfter);
                    }      
                    if (igColumn.lib4x?.spreadsheetViewOptions?.msoNumberFormat)
                    {
                        svColumn.lib4x.msoNumberFormat = igColumn.lib4x.spreadsheetViewOptions.msoNumberFormat;
                    }                               
                    // an IG column 'Always Read Only' will have item_type 'HIDDEN'
                    if (itemType == 'TEXT') 
                    {
                        svColumn.readOnly = false;
                        let maxlength = apexItem.element.attr('maxlength');
                        if (maxlength)
                        {
                            svColumn.lib4x.maxlength = maxlength;
                        }
                        let textCase = apexItem.element.attr('data-text-case');
                        if (textCase)
                        {
                            svColumn.lib4x.textCase = textCase;
                        } 
                        let trimSpaces = apexItem.element.attr('data-trim-spaces');
                        if (trimSpaces)
                        {
                            svColumn.lib4x.trimSpaces = trimSpaces;
                        }    
                        // any transformations to derive the value for wsData from the input value  
                        // where the input value might be entered or pasted                     
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            let result = newValue;
                            if (result)
                            {
                                if (trimSpaces)
                                {
                                    result = trimSpaces == 'LEADING' ? result.trimStart() : trimSpaces == 'TRAILING' ? result.trimEnd() : result;
                                    // effectively, when trimSpaces == 'NONE', no trimming
                                }   
                                else
                                {
                                    // default is both leading and trailing
                                    result = result.trim();
                                }                         
                                if (textCase)
                                {
                                    result = textCase == 'UPPER' ? result.toUpperCase() : textCase == 'LOWER' ? result.toLowerCase() : result;
                                }
                            }
                            return result;
                        }                                                
                    }
                    if (itemType == 'NUMBER') 
                    {
                        // contrary to IG model, numbers in the worksheet are kept as real numbers
                        svColumn.type = svEditors.numberInterface;
                        svColumn.lib4x.columnType = 'lib4x_number';
                        svColumn.readOnly = false;
                        svColumn.render = null;
                        svColumn.customFormat = igColumn.formatMask ?? null; 
                        let minimum = apexItem.element.data('min');
                        if (minimum != null)
                        {
                            svColumn.lib4x.minimum = minimum;
                        }
                        let maximum = apexItem.element.data('max');
                        if (maximum != null)
                        {
                            svColumn.lib4x.maximum = maximum;
                        }
                        // toDataValue: as stored in WorksheetOptions.data
                        svColumn.lib4x.toDataValue = function(value) {
                            // value '0' will be truthy (Boolean('0') evaluates to true)
                            // value '' will be falsy (Boolean('') evaluates to false)
                            // when the user is emptying a cell in a number column, JSS will give the value as empty string ''
                            let dataValue = '';
                            if (value != null)
                            {
                                dataValue = value && util.valueIsString(value) ? apex.locale.toNumber(value, svColumn.customFormat) : value;
                            }
                            return isNaN(dataValue) ? value : String(dataValue);
                        }
                        svColumn.lib4x.modelToDataValue = function(value) {
                            return svColumn.lib4x.toDataValue(value);
                        }                        
                        svColumn.lib4x.toFormattedValue = function(value) {
                            // behavior apex.locale.formatNumber:
                            // '' will stay '' 
                            // null will become ''
                            // NaN wil become 'NaN'
                            // undefined will give a 'not defined' error
                            let unformattedValue = value;
                            if (value && util.valueIsString(value))
                            {
                                unformattedValue = Number(value);
                            }
                            return isNaN(unformattedValue) ? value: apex.locale.formatNumber(unformattedValue, svColumn.customFormat).trim();
                        }
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            return svColumn.lib4x.toDataValue(newValue);
                        }                        
                        // toModelValue: as stored in IG model
                        svColumn.lib4x.toModelValue = svColumn.lib4x.toFormattedValue;
                        svColumn.render = function (td, value, x, y, instance, options) {
                            td.textContent = svColumn.lib4x.toFormattedValue(value);
                        }                            
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {   
                            let result = {textValue: '', htmlValue: ''};
                            let dataValue = instance.getValueFromCoords(x, y, false);                    
                            if (dataValue != null && dataValue !== '')
                            {
                                let formattedValue = svColumn.lib4x.toFormattedValue(dataValue);
                                result.textValue = formattedValue;
                                result.htmlValue = dataValue;
                                result.msoNumberFormat = getConfiguredMsoNumberFormat(svColumn, config.options.numberMsoNumberFormat) ?? 
                                                         util.oracleNumberMaskToMsoNumberFormat(svColumn.customFormat);   // will give 'General' when empty
                            }
                            return result;
                        }  
                        // data value to JavaScript native value
                        // returns a number value, null or NaN
                        svColumn.lib4x.dataToNativeValue = function(dataValue)
                        {
                            return (dataValue !== '' && dataValue != null) ? Number(dataValue) : null;
                        }      
                        // canonical value to data value  
                        // in case of null or Number.isNaN, '' is returned
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            let dataValue = '';
                            if (value != null && !Number.isNaN(value))
                            {
                                dataValue = String(value);
                            }
                            return dataValue;
                        }    
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.toFormattedValue(value);
                        }                                                                   
                    }
                    if (itemType == 'SINGLE_CHECKBOX') 
                    {
                        svColumn.readOnly = false;
                        svColumn.lib4x.columnType = 'lib4x_checkbox';
                        svColumn.render = null;
                        svColumn.lib4x.checkedValue = apexItem.element.attr('value') || 'Y';
                        svColumn.lib4x.uncheckedValue = apexItem.element.attr('data-unchecked-value') || 'N';
                        svColumn.lib4x.label = {};
                        svColumn.lib4x.label[svColumn.lib4x.checkedValue] = apex.lang.getMessage('APEX.ITEM_TYPE.CHECKBOX.CHECKED');
                        svColumn.lib4x.label[svColumn.lib4x.uncheckedValue] = apex.lang.getMessage('APEX.ITEM_TYPE.CHECKBOX.UNCHECKED');
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {   
                            let dataValue = instance.getValueFromCoords(x, y, false);
                            let label = svColumn.lib4x.label[dataValue] ?? '';
                            return {textValue: label, htmlValue: label};
                        };
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            if (newValue == svColumn.lib4x.label[svColumn.lib4x.checkedValue])
                            {
                                newValue = svColumn.lib4x.checkedValue;
                            }
                            else if (newValue == svColumn.lib4x.label[svColumn.lib4x.uncheckedValue])
                            {
                                newValue = svColumn.lib4x.uncheckedValue;
                            }
                            else if ((newValue != svColumn.lib4x.checkedValue) && (newValue != svColumn.lib4x.uncheckedValue))
                            {
                                // if non-valid value is pasted, assume unchecked
                                newValue = svColumn.lib4x.uncheckedValue;
                            }
                            return newValue;
                        }                                                
                        svColumn.type = svEditors.checkboxInterface;
                    }
                    if (itemType == 'CHECKBOX') // SWITCH!!
                    {
                        svColumn.readOnly = false;
                        svColumn.lib4x.columnType = 'lib4x_switch';
                        svColumn.render = null;
                        svColumn.lib4x.onValue = apexItem.element.attr('value') || 'Y';
                        svColumn.lib4x.offValue = apexItem.element.attr('data-off-value') || 'N';
                        svColumn.lib4x.label = {};
                        svColumn.lib4x.label[svColumn.lib4x.onValue] = apexItem.element.attr('data-on-label') || 'On';
                        svColumn.lib4x.label[svColumn.lib4x.offValue] = apexItem.element.attr('data-off-label') || 'Off';
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {   
                            let dataValue = instance.getValueFromCoords(x, y, false);
                            let label = svColumn.lib4x.label[dataValue.v] ?? '';
                            return {textValue: label, htmlValue: label};
                        };
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            let result = newValue;
                            // if newValue is composite, no action needed
                            if (newValue && util.valueIsString(newValue)) 
                            {
                                if (newValue == svColumn.lib4x.label[svColumn.lib4x.onValue])
                                {
                                    newValue = svColumn.lib4x.onValue;
                                }
                                else if (newValue == svColumn.lib4x.label[svColumn.lib4x.offValue])
                                {
                                    newValue = svColumn.lib4x.offValue;
                                }
                                else if ((newValue != svColumn.lib4x.onValue) && (newValue != svColumn.lib4x.offValue))
                                {
                                    // non-valid value is pasted
                                    newValue = '';
                                }
                                result = newValue ? {v: newValue, d: svColumn.lib4x.label[newValue]} : {v: '', d: ''};
                            }
                            return result;
                        }   
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.label[value] ?? '';
                        }  
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            let result = '';
                            if (value === svColumn.lib4x.onValue || value === svColumn.lib4x.offValue)
                            {
                                result = {v: value, d: svColumn.lib4x.label[value]};
                            }
                            return result;
                        }                                                                                            
                        svColumn.type = svEditors.switchInterface;
                    }         
                    if ((itemType == 'RADIO_GROUP') && (apexItem.element.find('.apex-item-option').length == 2))
                    {
                        svColumn.readOnly = false;
                        svColumn.lib4x.columnType = 'lib4x_simple_radio';
                        // reserve a bit more space
                        svColumn.width = svColumn.width * 1.5;
                        svColumn.render = null;
                        let firstRadio = apexItem.element.find('.apex-item-option').eq(0).find('input[type="radio"]');
                        let secondRadio = apexItem.element.find('.apex-item-option').eq(1).find('input[type="radio"]')
                        svColumn.lib4x.firstValue = firstRadio.attr('value');
                        svColumn.lib4x.secondValue = secondRadio.attr('value');
                        svColumn.lib4x.label = {};
                        svColumn.lib4x.label[svColumn.lib4x.firstValue] = firstRadio.attr('data-display');
                        svColumn.lib4x.label[svColumn.lib4x.secondValue] = secondRadio.attr('data-display');
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {   
                            let dataValue = instance.getValueFromCoords(x, y, false);
                            let label = svColumn.lib4x.label[dataValue.v] ?? '';
                            return {textValue: label, htmlValue: label};
                        };
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            let result = newValue;
                            // if newValue is composite, no action needed
                            if (newValue && util.valueIsString(newValue)) 
                            {
                                if (newValue == svColumn.lib4x.label[svColumn.lib4x.firstValue])
                                {
                                    newValue = svColumn.lib4x.firstValue;
                                }
                                else if (newValue == svColumn.lib4x.label[svColumn.lib4x.secondValue])
                                {
                                    newValue = svColumn.lib4x.secondValue;
                                }
                                else if ((newValue != svColumn.lib4x.firstValue) && (newValue != svColumn.lib4x.secondValue))
                                {
                                    // non-valid value is pasted
                                    newValue = '';
                                }
                                result = newValue ? {v: newValue, d: svColumn.lib4x.label[newValue]} : {v: '', d: ''};
                            }
                            return result;
                        }      
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.label[value] ?? '';
                        }            
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            let result = '';
                            if (value === svColumn.lib4x.firstValue || value === svColumn.lib4x.secondValue)
                            {
                                result = {v: value, d: svColumn.lib4x.label[value]};
                            }
                            return result;
                        }                                                                                  
                        svColumn.type = svEditors.simpleRadioInterface;                        
                    }   
                    // switch as pill buttons 
                    if ((itemType != 'RADIO_GROUP') && (itemType != 'CHECKBOX') && (apexItem.element.is('.apex-button-group, apex-item-group--switch')))
                    {
                        svColumn.readOnly = false;
                        svColumn.lib4x.columnType = 'lib4x_pill_buttons';
                        // reserve a bit more space
                        svColumn.width = svColumn.width * 1.2;
                        svColumn.render = null;
                        let firstRadio = apexItem.element.find('.apex-item-option').eq(0).find('input[type="radio"]');
                        let secondRadio = apexItem.element.find('.apex-item-option').eq(1).find('input[type="radio"]')
                        svColumn.lib4x.firstValue = firstRadio.attr('value');
                        svColumn.lib4x.secondValue = secondRadio.attr('value');
                        svColumn.lib4x.label = {};
                        svColumn.lib4x.label[svColumn.lib4x.firstValue] = firstRadio.parent().find('label').text();
                        svColumn.lib4x.label[svColumn.lib4x.secondValue] = secondRadio.parent().find('label').text();
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {   
                            let dataValue = instance.getValueFromCoords(x, y, false);
                            let label = svColumn.lib4x.label[dataValue.v] ?? '';
                            return {textValue: label, htmlValue: label};
                        };
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            let result = newValue;
                            // if newValue is composite, no action needed
                            if (newValue && util.valueIsString(newValue)) 
                            {
                                if (newValue == svColumn.lib4x.label[svColumn.lib4x.firstValue])
                                {
                                    newValue = svColumn.lib4x.firstValue;
                                }
                                else if (newValue == svColumn.lib4x.label[svColumn.lib4x.secondValue])
                                {
                                    newValue = svColumn.lib4x.secondValue;
                                }
                                else if ((newValue != svColumn.lib4x.firstValue) && (newValue != svColumn.lib4x.secondValue))
                                {
                                    // non-valid value is pasted
                                    newValue = '';
                                }
                                result = newValue ? {v: newValue, d: svColumn.lib4x.label[newValue]} : {v: '', d: ''};
                            }
                            return result;
                        } 
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.label[value] ?? '';
                        }     
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            let result = '';
                            if (value === svColumn.lib4x.firstValue || value === svColumn.lib4x.secondValue)
                            {
                                result = {v: value, d: svColumn.lib4x.label[value]};
                            }
                            return result;
                        }                                                                                           
                        svColumn.type = svEditors.pillButtonsInterface;                        
                    }                      
                    if (itemType == 'DATE_PICKER')
                    {
                        svColumn.readOnly = false;
                        svColumn.lib4x.columnType = 'lib4x_date_picker'
                        svColumn.render = null;
                        let dateFormat = apexItem.element.attr('format');  // regular attribute!
                        if (!dateFormat)
                        {
                            dateFormat = apex.locale.getDateFormat();
                        }                             
                        svColumn.customFormat = dateFormat;
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {
                            let result = {textValue: '', htmlValue: ''};
                            let dataValue = instance.getValueFromCoords(x, y, false); 
                            if (dataValue)
                            {
                                let customFormat = svColumn.customFormat;
                                let inclTime = dateFormat.includes('HH');
                                result.textValue = dataValue;
                                try
                                {
                                    let dateValue = apex.date.parse(dataValue, customFormat);
                                    result.htmlValue = jssDateToEpochUTC(dateValue);
                                }
                                catch(e)
                                {
                                    // if not able to parse, just put value as entered
                                    result.htmlValue = dataValue;
                                }
                                // used before:
                                // let msoNumberFormat = 'yyyy-mm-dd' + (inclTime ? ' hh:mm:ss' : '');
                                // let msoNumberFormat = inclTime ? 'General Date' : dateFormat.includes('MON') ? 'Medium Date' : 'Short Date';
                                // note: Medium Date results in Excel in year with 2 digits only
                                result.msoNumberFormat = getConfiguredMsoNumberFormat(svColumn, config.options.dateMsoNumberFormat) ?? oracleDateMaskToMsoNumberFormat(dateFormat);
                            }
                            return result;
                        }    
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            newValue = dateInputToDataValue(svColumn.customFormat, newValue);
                            return newValue;
                        } 
                        // data value to JavaScript native value
                        // returns a date object or null
                        svColumn.lib4x.dataToNativeValue = function(dataValue)
                        {
                            let nativeValue = null;
                            if (dataValue !== '' && dataValue != null)
                            {
                                try
                                {
                                    nativeValue = apex.date.parse(dataValue, svColumn.customFormat);
                                }
                                catch(e) {} 
                            }                          
                            return nativeValue;
                        }  
                        // canonical value to data value  
                        // in case of null or isNaN (Invalid Date) '' is returned
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            let dataValue = '';
                            if (value != null && !isNaN(value) && value instanceof Date)
                            {
                                try
                                {
                                    dataValue = apex.date.format(value, svColumn.customFormat);
                                }
                                catch(e){}
                            }
                            return dataValue;
                        } 
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.valueToDataValue(value);
                        }                                                                                           
                        svColumn.type = svEditors.dateInterface;                   
                    }
                    if ((itemType == 'SELECT') || ((itemType == 'RADIO_GROUP') && (apexItem.element.find('.apex-item-option').length != 2)))
                    {
                        // in case the IG column has 'Cascading List of Values' configured, we can not support editing in JSS
                        // when 'Cascading List of Values', the apex select item has the reinit method implemented
                        svColumn.readOnly = (apexItem.hasOwnProperty('reinit'));
                        svColumn.lib4x.columnType = 'lib4x_select_list';
                        // below will set lovMap and lovHasDuplicateDisplayValues
                        deriveLovMapFromItem(svColumn);                     
                        svColumn.lib4x.toDataValue = function(value) {
                            let result = value;
                            // when value is composite, no action needed
                            // if string, seek value and create composite
                            if (value && util.valueIsString(value)) {
                                let displayValue = util.normalizeDisplayValue(value);
                                let optionValue = svColumn.lib4x.lovMap.get(displayValue);
                                if (optionValue) {
                                    result = {v: optionValue, d: displayValue};
                                }
                                else
                                {
                                    result = {v: '', d: ''};
                                }
                            }
                            return result;
                        }
                        // model value to JSS dataset value
                        svColumn.lib4x.modelToDataValue = function(value) {
                            return svColumn.lib4x.toDataValue(value);
                        }
                        // user manual input value (display value) to JSS dataset value
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            return svColumn.lib4x.toDataValue(newValue);
                        }
                        // programmatic value (return value) to JSS dataset value
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            return util.lovMapValueForReturnValue(svColumn.lib4x.lovMap, value);                           
                        }
                        svColumn.lib4x.toClipboardValue = function(instance, x, y) {
                            let result = {textValue: '', htmlValue: ''};
                            let dataValue = instance.getValueFromCoords(x, y, false); 
                            if (dataValue)
                            {
                                result = {textValue: dataValue.d, htmlValue: dataValue.d};
                            }
                            return result;
                        }  
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.valueToDataValue(value)?.d;
                        }                                               
                        svColumn.render = null;
                        svColumn.type = svEditors.selectListInterface;
                    }
                    if (itemType == 'COLOR_PICKER') 
                    {
                        // just use the JSS color picker - is also able to process hex/rbg
                        svColumn.type = 'color';
                        svColumn.readOnly = false;
                    }
                    // Popup LOV, single value (in SV, a Select One will be used)
                    if (apexItem.element.hasClass('apex-item-popup-lov') && (apexItem.storageType == null))
                    {
                        // in case the IG column has 'Cascading List of Values' configured, we can not support editing in JSS
                        // when 'Cascading List of Values', the apex select item has the reinit method implemented
                        svColumn.readOnly = (apexItem.hasOwnProperty('reinit'));
                        svColumn.lib4x.columnType = 'lib4x_lov_one';
                        svColumn.render = null;
                        svColumn.lib4x.lovMap = new Map();
                        svColumn.lib4x.lovHasDuplicateDisplayValues = false;
                        svColumn.lib4x.toDataValue = function(value) {
                            let result = value;
                            // when value is composite, no action needed
                            // if string, seek value and create composite
                            if (value && util.valueIsString(value)) {
                                let returnValue = svColumn.lib4x.lovMap.get(util.normalizeDisplayValue(value));
                                if (returnValue) {
                                    result = {v: returnValue, d: value};
                                }
                                else
                                {
                                    result = {v: '', d: ''};
                                }
                            }
                            return result;
                        }
                        svColumn.lib4x.modelToDataValue = function(value) {
                            return svColumn.lib4x.toDataValue(value);
                        }
                        svColumn.lib4x.inputToDataValue = function(newValue) {
                            return svColumn.lib4x.toDataValue(newValue);
                        }  
                        svColumn.lib4x.valueToDataValue = function(value)
                        {
                            return util.lovMapValueForReturnValue(svColumn.lib4x.lovMap, value);                           
                        } 
                        svColumn.lib4x.displayValueFor = function(value)    
                        {
                            return svColumn.lib4x.valueToDataValue(value)?.d;
                        }                                                
                        svColumn.type = svEditors.selectOneInterface;	
                    }                    
                    // revert readOnly setting if needed
                    // in case IG is not editable, all IG columns will be readonly
                    // the IG grid has an allowEditMode setting, but that one only dictates if mouse/keyboard can be used to enter edit mode
                    // include check on is-readonly class (https://docs.oracle.com/en/database/oracle/apex/24.2/aexjs/grid.html#classes-section)
                    let columnHasIsReadOnlyClass = (igColumn.columnCssClasses && igColumn.columnCssClasses.split(/\s+/).includes('is-readonly'));
                    if (igColumn.readonly || columnHasIsReadOnlyClass || igColumn.lib4x?.spreadsheetViewOptions?.readOnly)
                    {
                        svColumn.readOnly = true;
                    }
                    else 
                    {
                        // for 'Cascading List of Values', the parent column should be read only
                        // parent will have an on change refresh handler (as to refresh the child)
                        let changeEvents = $._data(apexItem.node, "events")?.change;
                        if (changeEvents && (changeEvents.filter((evt)=>{return(evt.handler.name=='refresh')}).length > 0))
                        {
                            svColumn.readOnly = true;
                        }  
                    }                  
                    svColumns.push(svColumn);
                }
            }
            // compose all inputs as needed for createSpreadsheet
            let model = gridView.model;
            let y = 0;
            let wsMeta = null;
            let dsMeta = new Map();
            let wsData = [];
            let wsIds = [];
            let wsReadOnlyCells = new Set();
            let hasHighlights = false;
            model.forEach(function (record, index, recordId) {
                let recMeta = model.getRecordMetadata(recordId);
                if (!recMeta.deleted && !recMeta.agg && y < config.options.maxRows) {
                    let editAllowed = model.allowEdit(record);
                    let deleteAllowed = model.allowDelete(record);
                    let rowReadOnly = (!editAllowed);
                    let wsRow = [];
                    let rowMeta = dsMetaUtil.initRowMeta();
                    ['inserted', 'updated'].forEach(prop => {
                        if (prop in recMeta) {
                            rowMeta.gv[prop] = recMeta[prop];
                        }
                    });
                    rowMeta.gv.editAllowed = editAllowed;
                    rowMeta.gv.deleteAllowed = deleteAllowed;   
                    if (modelUtil.recordHasMessages(model, recordId, 'error'))
                    {
                        dsMetaUtil.addIssue(rowMeta, {
                            recordId: recordId,
                            // Grid row has validation error(s)
                            message: getMessage('GRID_ROW_HAS_ERROR')
                        });                       
                    }                   
                    if (modelUtil.recordHasMessages(model, recordId, 'warning'))
                    {
                        dsMetaUtil.addIssue(rowMeta, {
                            recordId: recordId,
                            //Grid row has warning message(s)
                            message: getMessage('GRID_ROW_HAS_WARNING')
                        });                       
                    }        
                    if (recMeta.error || recMeta.warning)
                    {
                        dsMetaUtil.addIssue(rowMeta, {
                            recordId: recordId,
                            message: recMeta.message
                        });                         
                    } 
                    if (config.options.applyHighlighting)
                    {
                        if (recMeta.highlight && !(isNaN(recMeta.highlight)))
                        {
                            rowMeta.gv.highlight = recMeta.highlight;
                            hasHighlights = true;
                        }  
                    }                             
                    svColumns.forEach((svColumn, x) => {
                        if (svColumn.name)
                        {
                            let svValue = model.getValue(record, svColumn.name);
                            if (svColumn.lib4x?.modelToDataValue)
                            {
                                svValue = svColumn.lib4x.modelToDataValue(svValue);
                            } 
                            wsRow.push(svValue);
                            let valueChanged = modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'changed');
                            if (valueChanged)
                            {   
                                dsMetaUtil.setFieldMeta(rowMeta, ORIG_GV, svColumn.name, 'changed', true);
                            }
                            if (config.options.applyHighlighting)
                            {
                                let highlight = modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'highlight');
                                // pick up only when number
                                if (highlight && !(isNaN(highlight)))
                                {   
                                    dsMetaUtil.setFieldMeta(rowMeta, ORIG_GV, svColumn.name, 'highlight', highlight);
                                    hasHighlights = true;
                                }    
                            }                        
                            let fieldHasError = modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'error');
                            if (fieldHasError)
                            {   
                                let message = modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'message');
                                dsMetaUtil.setFieldMeta(rowMeta, ORIG_GV, svColumn.name, 'error', fieldHasError);
                                dsMetaUtil.setFieldMeta(rowMeta, ORIG_GV, svColumn.name, 'message', message);
                            } 
                            else
                            {   
                                let fieldHasWarning = modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'warning');
                                if (fieldHasWarning)
                                {   
                                    let message = modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'message');
                                    dsMetaUtil.setFieldMeta(rowMeta, ORIG_GV, svColumn.name, 'warning', fieldHasWarning);
                                    dsMetaUtil.setFieldMeta(rowMeta, ORIG_GV, svColumn.name, 'message', message);
                                }  
                            }                                                    
                            if (!svColumn.readOnly)
                            {
                                let cellReadOnly = rowReadOnly;
                                if (!cellReadOnly && config.options.applyReadOnlyCells)
                                {
                                    cellReadOnly = (!!modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'ck')) ||
                                                   (!!modelUtil.getFieldMetaPropertyValue(recMeta, svColumn.name, 'disabled'));
                                }
                                if (cellReadOnly)
                                {
                                    let cellName = jspreadsheet.helpers.getCellNameFromCoords(x,y);
                                    wsReadOnlyCells.add(cellName);
                                }
                            }
                            if (svColumn.lib4x.columnType == 'lib4x_lov_one')
                            {
                                let lovMap = svColumn.lib4x.lovMap;
                                if (svValue.v && svValue.d)
                                {
                                    let displayValue = util.normalizeDisplayValue(svValue.d);
                                    let mappedValue = lovMap.get(displayValue);
                                    if ((mappedValue !== undefined) && (mappedValue != svValue.v))
                                    {
                                        svColumn.lib4x.lovHasDuplicateDisplayValues = true;
                                        // duplicate display values we can't handle as we won't be able to
                                        // determine unambiguously the return value
                                        svColumn.readOnly = true;
                                    }
                                    else
                                    {
                                        lovMap.set(displayValue, svValue.v);
                                    }
                                }
                            }
                        }
                    });
                    y = y + 1;
                    wsData.push(wsRow);
                    wsIds.push(recordId);
                    dsMeta.set(recordId, rowMeta);
                }          
            });
            if (y == 0 && model.allowAdd())
            {
                // JSS wants minimal 1 row!
                let wsRow = [];
                let rowMeta = dsMetaUtil.initRowMeta();
                rowMeta.sv.inserted = true;
                rowMeta.gv.editAllowed = true;
                rowMeta.gv.deleteAllowed = true;                
                let recordId = getTempRecordId();                
                svColumns.forEach((svColumn, x) => {
                    if (svColumn.name) 
                    {
                        let svValue = svColumn.lib4x.defaultValue || '';
                        if (svColumn.lib4x?.toDataValue)
                        {
                            svValue = svColumn.lib4x.toDataValue(svValue);
                        } 
                        wsRow.push(svValue);
                    }
                });                
                wsData.push(wsRow);
                wsIds.push(recordId);
                dsMeta.set(recordId, rowMeta);
                y = y + 1;
            }       
            let nestedHeaders = null;
            if (igColumnGroups)
            {
                nestedHeaders = deriveNestedHeaders(igColumnGroups, svColumns);
            }
            // footers
            let footers = null;
            if (hasAggregators(svAggregators))
            {
                let footerDefs = getFooterDefs(svAggregators);
                footers = footerDefs.length > 0 ? footerDefs.map(fd => [fd.label]) : null;
            }
            createSpreadsheet(svColumns, y, wsData, wsIds, wsMeta, dsMeta, wsReadOnlyCells, svAggregators, footers, nestedHeaders, hasHighlights); 
        }   // end of createSpreadsheetView

        /*
         * Open the spreadsheet view (inline dialog)
        */
        function openSpreadsheetView(svStaticId, igStaticId) {  
            apex.message.hidePageSuccess();       
            let svStaticIdSv = svStaticId + SV_EXT;
            sv_igStaticId[svStaticIdSv] = igStaticId;  
            let model = apex.region(igStaticId).call('getViews').grid.model;
            if (!modelUtil.sparselyLoaded(model))
            {
                let spinner$ = null;
                // check if model has significant data
                // strictly speaking, also the pagination type would be relevant as for 
                // page pagination, no spinner would be required
                // using model.getTotalRecords(true) didn't work out as it gives the total from DB 
                // in case 'Show Total Count' is checked on the IG
                if (model._data.length >= 1000)
                {
                    spinner$ = apex.util.showSpinner($('#' + igStaticId));
                }
                setTimeout(()=>{
                    try {
                        createSpreadsheetView(svStaticId, false);
                    }
                    finally {
                        if (spinner$)
                        {
                            spinner$.remove();
                        }
                    }
                    openSvDialog(svStaticId);
                });
            }
            else
            {
                // a sparsely loaded model has gabs in the data
                // this can come into existence when the user is jumping for example from first page
                // to the last page (page pagination) or using scrollbar to quickly jump from top to bottom (scroll pagination)
                if (!model.isChanged())
                {
                    // no changes so we can just refresh to get rid of the gabs
                    let viewId = model.subscribe({
                        onChange: function(changeType, change) {
                            if (changeType == 'addData')
                            {
                                model.unSubscribe(viewId);
                                if (!getJWorksheet(svStaticIdSv))    // checking just to be sure
                                {
                                    openSpreadsheetView(svStaticId, igStaticId);
                                }
                            }
                        }
                    });
                    model.clearData();
                }
                else
                {
                    // proves out we can not do a model.fetchAll in situation of 
                    // a sparsely loaded model - it can result in duplicate data in the model (APEX bug?)
                    // 'Not able to open the Spreadsheet View. Please save and refresh the data and try again.'
                    apex.message.alert(getMessage('NOT_ABLE_TO_OPEN'));
                }
            }        
        }

        // opens the SV dialog region 
        function openSvDialog(svStaticId) {
            let svStaticIdSv = svStaticId + SV_EXT;
            let igStaticId = sv_igStaticId[svStaticIdSv];      
            $('#'+ svStaticId).closest('.ui-dialog').find('.ui-dialog-title').text(getWorksheetName(igStaticId));
            apex.theme.openRegion(svStaticId);
        }

        function getWorksheetName(igStaticId)
        {
            let igTitle = util.ig.getTitle(igStaticId); 
            // Sheet
            return igTitle + ' ' + getMessage('SHEET');
        }

        /*
         * Init the SV: setting up eventhandlers and an actions context. The real SV only get's created
         * upon opening the SV from the IG
         */
        let initSV = function (svStaticId) {
            let svStaticIdSv = svStaticId + SV_EXT;            
            let config = svConfig[svStaticIdSv];
            // SV only supported in (modal) dialog
            if (!($('#' + svStaticId).is(':ui-dialog'))) {
                throw new Error('IG Spreadsheet View region should be an inline Dialog or Drawer (' + svStaticId + ')');
            }
            let dlg$ = $('#' + svStaticId);
            dlg$.css('overflow', 'hidden');
            dlg$.find('.t-DialogRegion-bodyWrapperOut, .t-DrawerRegion-bodyWrapperOut').css('overflow', 'hidden');
            dlg$.dialog('option', {
                closeOnEscape: true
            });  
            let maxButton$ = null;
            let isDialogRegion = dlg$.hasClass('t-DialogRegion');  
            if (isDialogRegion)
            {
                // add maximize button
                let dlgWidget$ = dlg$.dialog("widget");  // will be the dialog wrapper
                maxButton$ = $("<button>", {
                    class: "ui-dialog-titlebar-max",
                    type: "button",
                    title: getMessage('DIALOG.MAXIMIZE')
                }).button({
                    icon: "ui-icon-max-restore",
                    showLabel: false
                });
                $(".ui-dialog-titlebar .ui-dialog-title", dlgWidget$).after(maxButton$);
                // on click event handler
                let originalState = null;
                function maximizeDialog() {
                    dlg$.dialog("option", {
                        width: Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0),
                        height: Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0),
                        position: {
                            my: "left top",
                            at: "left top",
                            of: window
                        }
                    });
                    dlg$.trigger('dialogresize');
                }
                maxButton$.on('click', function (e) {
                    let maximized = maxButton$.hasClass(C_IS_MAX);
                    if (!maximized) {
                        // Save current dialog size and position before maximizing
                        originalState = {
                            width: dlg$.dialog("option", "width"),
                            height: dlg$.dialog("option", "height"),
                            position: dlg$.dialog("option", "position")
                        };
                        maximizeDialog();
                        // Restore
                        maxButton$.addClass(C_IS_MAX).attr('title', getMessage('DIALOG.RESTORE'));
                    } else {
                        // restore original size and position
                        // take into account the browser window might be smaller now due to resizing
                        if (originalState) {
                            dlg$.find('.lib4x-SV .jss_content').width('auto');
                            const maxW = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
                            const maxH = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
                            const restoreWidth = Number.isFinite(originalState.width) ? Math.min(originalState.width, maxW - 20) : originalState.width;                            
                            const restoreHeight = Number.isFinite(originalState.height) ? Math.min(originalState.height, maxH - 20) : originalState.height;   
                            dlg$.dialog("option", {
                                width: restoreWidth,
                                height: restoreHeight,
                                position: originalState.position
                            });
                            dlg$.trigger('dialogresize');
                        }
                        maxButton$.removeClass(C_IS_MAX).attr('title', getMessage('DIALOG.MAXIMIZE'));
                    }
                });
                $(window).on('resize', function () {
                    if (dlg$.is(':visible'))
                    {
                        let maximized = maxButton$.hasClass(C_IS_MAX);
                        if (maximized) {
                            maximizeDialog();
                        }
                    }
                });                     
            }
            // Add OK and Cancel buttons to dialog bottom
            let rButtonsRregion$ = dlg$.find('.t-DialogRegion-buttons, .t-DrawerRegion-buttons').find(' .t-ButtonRegion-col--right .t-ButtonRegion-buttons');
            if (rButtonsRregion$.length)
            {
                let cancelBtnId = 'btn_' + svStaticId + '_cancel';
                let okBtnId = 'btn_' + svStaticId + '_ok';
                let okLabel = getMessage('DIALOG.OK');
                let cancelLabel = getMessage('DIALOG.CANCEL');
                rButtonsRregion$.append('<button class="t-Button" type="button" id="' + cancelBtnId + '"><span class="t-Button-label">' + cancelLabel + '</span></button>' +
                    '<button class="t-Button t-Button--hot" type="button" id="' + okBtnId + '"><span class="t-Button-label">' + okLabel + '</span></button>');
                $('#' + cancelBtnId).on('click', function(event){
                    apex.theme.closeRegion(svStaticId);
                })
                $('#' + okBtnId).on('click', function(event){
                    if (dataHasChanged(svStaticId))
                    {
                        syncToModel(svStaticId);
                        let igStaticId = sv_igStaticId[svStaticIdSv];
                        let ig$ = $('#' + igStaticId);
                        let wait = false;
                        let finalize = function(showMessage) {
                            if (showMessage)
                            {
                                showSyncResultMessage(igStaticId);
                            }
                        };                        
                        apex.event.trigger(ig$, 'lib4xspreadsheetviewsynchronize', {
                            igStaticId: igStaticId,
                            svStaticId: svStaticId,
                            registerAsync: function(promise) {
                                wait = true;
                                let showMessage = true;
                                promise.catch(() => {
                                    showMessage = false;
                                }).finally(() => {
                                    finalize(showMessage);
                                });
                            }                            
                        });
                        if (!wait)
                        {
                            finalize(true);
                        }
                    }
                    // at this point, there should be no open changes
                    // to be sure, set lib4x_okClose to true
                    dlg$.dialog('option', 'lib4x_okClose', true);   
                    apex.theme.closeRegion(svStaticId);
                });
            }   
            // prepare dialog resize event handler
            dlg$.on('dialogresize',
                apex.util.debounce(() => {
                    let jssContent$ = dlg$.find('.lib4x-SV .jss_content');
                    if (jssContent$.length)
                    {
                        let body$ = dlg$.find('.t-DialogRegion-body, .t-DrawerRegion-body');
                        let bodyWrapperOut$ = dlg$.find('.t-DialogRegion-bodyWrapperOut, .t-DrawerRegion-bodyWrapperOut');
                        let hPadding = body$.innerHeight() - body$.height();
                        let wPadding = body$.innerWidth() - body$.width();
                        let toolbarHeight = dlg$.find('.' + C_LIB4X_SV_TOOLBAR).outerHeight(true);   
                        let paginationbarHeight = dlg$.find('.jss_pagination:visible').outerHeight(true) || 0;
                        jssContent$.css({ "maxHeight": bodyWrapperOut$.height() - hPadding - toolbarHeight - paginationbarHeight - 3 + "px", "width": "auto", "maxWidth": dlg$.width() - wPadding - 2 + "px" });
                    }                          
                }, 100)
            );
            // prepare dialog open event handler
            dlg$.on('dialogopen', function (event) {            
                dlg$.dialog('option', 'lib4x_okClose', false);
                // upon using date picker popup in SV date cell, upon clicking things like 'Next Month', 
                // it shouldn't trigger JSS to call closeEditor() method
                // so we block the event from reaching the jss event listener
                // same concept applies to comboselect popup as used by select one
                $(document).off('mousedown.svEditor').on('mousedown.svEditor', '.a-DatePicker, .a-ComboSelect-popup', function (e) {
                    if ($('.' + C_LIB4X_SV + ' .' + C_SV_EDITOR).length) {
                        e.stopImmediatePropagation();
                    }
                });  
                $(window).on('resize.lib4x_sv', function () {
                    setTimeout(() => {
                        $('#' + svStaticId + ':visible').trigger('dialogresize');
                    }, 100);
                }); 
                $('#'+svStaticIdSv).on('keydown.lib4x_sv', function (e) {
                    if (e.key === 'Tab' && !e.ctrlKey) 
                    {
                        // support tabbing to next row in case giving tab in last column
                        let instance = getJWorksheet(svStaticIdSv);
                        if (instance)
                        {                     
                            let selectedCells = instance.getSelected(false);
                            if (selectedCells && selectedCells.length == 1) 
                            {
                                let x,y;
                                if (!e.shiftKey && (selectedCells[0].x == (instance.options.columns.length-1)))
                                {
                                    x = 0;
                                    // select first cell on next row; JSS will guard last row
                                    y = getNextVisibleRow(instance, selectedCells[0].x, selectedCells[0].y);
                                }
                                else if (e.shiftKey && (selectedCells[0].x == 0))
                                {
                                    x = instance.options.columns.length-1;
                                    // select last cell on previous row; JSS will guard first row
                                    y = getPreviousVisibleRow(instance, selectedCells[0].x, selectedCells[0].y);
                                }
                                if (x != null && y != null && y != selectedCells[0].y)                                   
                                {
                                    if (instance.edition != null)
                                    {
                                        let cell = selectedCells[0].element;
                                        if ($(cell).hasClass(C_EDITOR))
                                        {
                                            instance.closeEditor(cell, true);
                                        }
                                    }
                                    instance.updateSelectionFromCoords(x, y);
                                    e.preventDefault();
                                    e.stopPropagation(); 
                                }
                            }  
                        }                         
                    }
                    if (e.key !== 'Escape') return;
                    let isEditing = $('#'+svStaticIdSv).find('.editor').length > 0;
                    if (isEditing) 
                    {
                        // prevent dialog close (editor will be closed)
                        e.preventDefault();                           
                    }
                });                                                      
            });         
            dlg$.on('dialogbeforeclose', function (event) {
                function cleanUp()
                {
                    tempRecordSeq = 0;
                    lastClosedIgStaticId = sv_igStaticId[svStaticIdSv];
                    delete sv_igStaticId[svStaticIdSv];
                    // reset edit on focus mode / filter changes flag on actions
                    let actionsContext = apex.actions.findContext('IGSpreadsheetView', $('#' + svStaticIdSv)[0]);
                    let toggleEditOnFocusAction = actionsContext.lookup('toggle-edit-on-focus');
                    if (toggleEditOnFocusAction)
                    {
                        toggleEditOnFocusAction.editOnFocus = false;
                    }                    
                    let filterAction = actionsContext.lookup('radiogroup-filter-rows');
                    if (filterAction)
                    {
                        filterAction.filterChoice = FILTER_ALL;
                    }
                    destroySpreadsheet(svStaticId);             
                }
                
                let okClose = dlg$.dialog('option', 'lib4x_okClose');
                if (!okClose && dataHasChanged(svStaticId)) {
                    // 'New changes will be lost. Close the dialog?'
                    apex.message.confirm(getMessage('DIALOG.OPEN_CHANGES_Q_CLOSE'), function(okPressed) {
                        if (okPressed) 
                        {
                            cleanUp();
                            dlg$.dialog('option', 'lib4x_okClose', true);
                            dlg$.dialog('close');
                        }
                    });
                    event.preventDefault();
                }
                else
                {
                    cleanUp();
                }                
            });
            dlg$.on('dialogclose', function (event) {
                $(document).off('mousedown.svEditor');
                $(document).off('keydown.lib4x_sv');
                $(window).off('resize.lib4x_sv');
            });
            // preserve clipboard content upon paste
            // any further processing will be done in onbeforepaste event handler
            $('#'+svStaticIdSv).on('paste', function (jQueryEvent) {
                let instance = getJWorksheet(svStaticIdSv);
                if (instance)
                {
                    instance.options.lib4x.clipboardHtml = jQueryEvent.originalEvent?.clipboardData?.getData('text/html');
                }
            });              
            // trigger a resize when SV becomes visible 
            apex.widget.util.onVisibilityChange($('#' + svStaticIdSv), function (visible) {
                if (visible) {
                    setTimeout(() => {
                        dlg$.trigger('dialogresize');
                    }, 100);
                }
            });
            // set up actions context for SV widget
            let actionsContext = apex.actions.createContext('IGSpreadsheetView', $("#" + svStaticIdSv)[0]);
            actionsContext.add([
                { 
                    name: 'undo', 
                    icon: 'fa fa-undo',    
                    title: getMessage('UNDO_LAST_CHANGE'),      // Undo the last change 
                    action: function() {
                        let instance = getJWorksheet(svStaticIdSv);                    
                        instance.undo();
                    }
                },
                { 
                    name: 'redo',    
                    icon: 'fa fa-repeat',    
                    title: getMessage('REDO_RECENT_CHANGE'),    // Redo the most recent change       
                    action: function() {
                        let instance = getJWorksheet(svStaticIdSv);                        
                        instance.redo();
                    }
                },
                { 
                    name: 'add-row-before', 
                    labelKey: getMessage('ADD_BEFORE'),  // Add Before
                    title: getMessage('ADD_ROW_BEFORE'),      // Add a new row before              
                    action: function() {
                        let instance = getJWorksheet(svStaticIdSv);                       
                        let selection = instance.getSelection();
                        let rowIndex = selection ? selection[1] : 0;
                        instance.insertRow(1, rowIndex, true); // one row, before row rowIndex
                        if (selection == null)
                        {
                            // scroll to top
                            $('#' + svStaticIdSv + ' .jss_content').scrollTop(0);
                        }
                        instance.updateSelectionFromCoords(0, rowIndex, 0, rowIndex);                            
                    }
                },
                { 
                    name: 'add-row-after', 
                    label: getMessage('ADD_AFTER'),         // Add After
                    title: getMessage('ADD_ROW_AFTER'),     // Add a new row after                 
                    action: function() {
                        let instance = getJWorksheet(svStaticIdSv);                        
                        let selection = instance.getSelection();
                        selection ? instance.insertRow(1, selection[3], false) : instance.insertRow();
                        let rowIndex;
                        if (selection == null)
                        {
                            // scroll to bottom
                            $('#' + svStaticIdSv + ' .jss_content').scrollTop($('#' + svStaticIdSv + ' .jss_content')[0].scrollHeight);
                            rowIndex = instance.rows.length;
                        }
                        else
                        {
                            rowIndex = selection[3]+1;
                        }
                        instance.updateSelectionFromCoords(0, rowIndex, 0, rowIndex);
                    }
                },                  
                { 
                    name: 'delete-rows', 
                    label: getMessage('DELETE_ROWS'),  // Delete Row(s)
                    // can't use shortcut 'Delete' as the user won't be able anymore to delete cell content
                    action: function() {
                        let instance = getJWorksheet(svStaticIdSv);
                        let selection = instance.getSelectedRows();
                        if (selection && selection.length > 0)
                        {
                            if (instance.rows.length == 1)
                            {
                                apex.message.alert(getMessage('DELETE_LAST_ROW_NOT_POSSIBLE'));     // It is not possible to delete the last row
                            }
                            else
                            {
                                instance.deleteRow();   // will take current selection                            
                            }
                        }
                    }
                },                
                { 
                    name: 'load-all', 
                    label: getMessage('LOAD_ALL'),  // Load All
                    icon: 'fa fa-download-alt',                                              
                    action: function() {               
                        if (!dataHasChanged(svStaticId))
                        {
                            setTimeout(()=>{                            
                                loadAll(svStaticId);
                            });                                
                        }
                        else
                        {
                            // There are changes. Pls synchronize first.
                            apex.message.alert(getMessage('SYNCHRONIZE_FIRST'), null, {style: 'information'});      
                        }
                    }
                },
                { 
                    name: 'synchronize', 
                    label: getMessage('SYNCHRONIZE_WITH_GRID'),  // Synchronize with Grid
                    icon: 'fa fa-exchange',                                              
                    action: function() {
                        if (dataHasChanged(svStaticId))
                        {
                            // use timeout so editor gets closed first before existing spreadsheet is getting destroyed
                            // without, a-select gives 'cannot call methods on popup prior to initialization'
                            setTimeout(()=>{  
                                let svOverlay$ = setDialogOverlay(svStaticId, true);
                                syncToModel(svStaticId);                                
                                let svStaticIdSv = svStaticId + SV_EXT;  
                                let igStaticId = sv_igStaticId[svStaticIdSv];
                                let ig$ = $('#' + igStaticId);
                                let wait = false;
                                let finalize = function(showMessage) {
                                    if ($('#' + svStaticId).is(':visible'))
                                    {
                                        createSpreadsheetView(svStaticId, true);
                                    }
                                    svOverlay$.remove();    
                                    if (showMessage)
                                    {
                                        showSyncResultMessage(igStaticId);
                                    }
                                };
                                apex.event.trigger(ig$, 'lib4xspreadsheetviewsynchronize', {
                                    igStaticId: igStaticId,
                                    svStaticId: svStaticId,
                                    loadingIndicator: '#' + svStaticIdSv,
                                    registerAsync: function(promise) {
                                        wait = true;
                                        let showMessage = true;
                                        promise.catch(() => {
                                            // showing success message would suppress existing error message
                                            showMessage = false;                                          
                                        }).finally(() => {
                                            finalize(showMessage);
                                        });
                                    }
                                });
                                if (!wait) {
                                    finalize(true);
                                }                                                    
                            });
                        }
                    }
                },                
                {
                    name: 'toggle-edit-on-focus',
                    //label: 'Edit on Focus',
                    title: getMessage('EDIT_ON_FOCUS'),     // Edit on Focus
                    editOnFocus: false,
                    icon: 'fa fa-i-cursor',
                    set: function(editOnFocus) 
                    {
                        this.editOnFocus = editOnFocus;  
                        setTimeout(()=>{
                            $('#'+svStaticIdSv).focus();  
                        }, 10);
                    },
                    get: function()
                    {
                        return this.editOnFocus;
                    }
                },
                {
                    name: 'radiogroup-filter-rows',
                    filterChoice: FILTER_ALL,
                    set: function(filterChoice) 
                    {
                        this.filterChoice = filterChoice;
                        let jssContent$ = dlg$.find('.lib4x-SV .jss_content');
                        jssContent$.scrollTop(0);                         
                        setRowsVisibiliy(svStaticIdSv);
                        setTimeout(()=>{
                            $('#'+svStaticIdSv).focus();  
                        }, 10);                        
                    },
                    get: function()
                    {
                        return this.filterChoice;
                    },
                    choices: [
                        {
                            label: getMessage('SHOW_ALL'),  // Show all
                            value: FILTER_ALL,
                            //icon: 'fa fa-list-alt'
                            icon: 'fa fa-table'
                        },
                        {
                            label: getMessage('SHOW_MODIFIED_ROWS'),    // Show modified rows
                            value: FILTER_MODIFIED,
                            //icon: 'fa fa-list-alt fam-star fam-is-disabled'
                            icon: 'fa fa-table-new'
                        },
                        {
                            label: getMessage('SHOW_ISSUES_ROWS'),     // Show rows with issues
                            value: FILTER_ISSUES,
                            //icon: 'fa fa-list-alt fam-warning fam-is-disabled'
                            icon: 'fa fa-table-x'
                        } 
                    ]
                },      
                { 
                    name: 'switch-pagination-type', 
                    label: getMessage('SWITCH_PAGINATION'),  // Switch Scroll/Page Pagination                            
                    action: function() {               
                        if (!dataHasChanged(svStaticId))
                        {
                            setTimeout(()=>{                            
                                switchPaginationType(svStaticId);
                            });                                
                        }
                        else
                        {
                            // There are changes. Pls synchronize first.
                            apex.message.alert(getMessage('SYNCHRONIZE_FIRST'), null, {style: 'information'});      
                        }
                    }
                },                        
                {
                    name: 'toggle-highlighting',
                    title: getMessage('TOGGLE_HIGHLIGHTING'),       // Toggle Highlighting
                    showHighlight: true,  // initial state of the toggle
                    icon: 'fa fa-star-o',
                    set: function(showHighlight) 
                    {
                        this.showHighlight = showHighlight;  
                        // to toggle highlighting, CSS proved not effective; instead we rename the hlr_*/hlc_* classes
                        if (showHighlight)
                        {
                            $('#'+svStaticIdSv).find('[class*="hlr_off_"], [class*="hlc_off_"]').each(function () {
                                this.className = this.className.replace(/\bhlr_off_(\d+)\b/g, 'hlr_$1').replace(/\bhlc_off_(\d+)\b/g, 'hlc_$1');
                            });
                        }
                        else
                        {
                            $('#'+svStaticIdSv).find('[class*="hlr_"], [class*="hlc_"]').each(function () {
                                this.className = this.className.replace(/\bhlr_(\d+)\b/g, 'hlr_off_$1').replace(/\bhlc_(\d+)\b/g, 'hlc_off_$1');
                            });                            
                        }
                        setTimeout(()=>{
                            $('#'+svStaticIdSv).focus();  
                        }, 10);
                    },
                    get: function()
                    {
                        return this.showHighlight;
                    }
                },  
                { 
                    name: 'show-help', 
                    label: getMessage('HELP'),      // Help
                    icon: 'fa fa-question-circle-o',                                              
                    action: function() {
                        let helpText = config.options.helpText ||
                        '<p>The Spreadsheet View loads with a copy of the Grid data. You can edit it, add/delete rows, and copy/paste from and to Excel. ' +
                        'Upon \'OK\', the changes are synchronized back to the Grid. In the Grid, you can address any issues like validation errors and then save the data.</p>' +
                        '<p>Upon selecting a cell, you can directly start typing to replace any current value. Or double-click or use <u>F2</u> to change the existing value.</p>' +
                        '<p>You can use the other familiar spreadsheet type of editing features like selecting cell(s), copy them and paste the values elsewhere. Or use the '+
                        'Fill Handle (bottom-right corner of a selection) to copy or fill into adjacent cells. You can also use copy-and-paste to/from Excel.<p/>' +
                        '<p><u>Edit on Focus</u>: this button enables you to edit cells without need to first use F2 or double click.</p>' +
                        '<p><u>Load All</u>: initially, a subset of data might have been loaded only. Use this button to load all the data. It loads to a maximum of ' + config.options.maxRows + ' rows.</p>' +
                        '<p><u>Synchronize</u>: this button lets you synchronize your changes in between with the Grid without closing the dialog. Any resulting validation errors will be marked and shown in the spreadsheet.</p>' +
                        '<p><u>Ctrl+Z/Ctrl+Y</u>: shortcut keys for Undo/Redo. This will apply to changes which haven\'t been synchronized to the Grid yet.</p>' +
                        (config.options.additionalHelpText || '');
                        // Interactive Grid Spreadsheet View
                        apex.theme.popupFieldHelp({title: getMessage('HELP_TITLE'), helpText: helpText});
                    }
                }                                       
            ]);
        }

        function getEventHandler(igStaticId, handlerName)
        {
            let result = null;
            if (sv_eventHandlers.hasOwnProperty(igStaticId))
            {
                let eventHandlers = sv_eventHandlers[igStaticId];
                if (typeof eventHandlers[handlerName] === 'function')     
                {
                    result = eventHandlers[handlerName];
                }   
            }
            return result;    
        }

        function fireSVEvent(igStaticId, name, ctx, instance, ignoreHistory)
        {
            if (ignoreHistory && instance) {
                instance.ignoreHistory = true;
            }
            getEventHandler(igStaticId, name)?.(ctx);
            if (ignoreHistory && instance) {
                instance.ignoreHistory = false;
            }
        }

        function getWorksheetInterface(instance)
        {
            let columnsByName = instance.options.lib4x.columnsByName;                                        
            let worksheetInterface = {
                // order of arguments in line with the model API (recordId, fieldName)
                // which is opposite to how JSS is doing it (colIndex, rowIndex)
                getValue: function(rowIndex, columnName)
                {
                    let result = null;
                    let svColumn = columnsByName[columnName];
                    if (svColumn)
                    {
                        let colIndex = svColumn.index;
                        if (colIndex >= 0)
                        {
                            result = toCanonicalValue(svColumn, instance.getValueFromCoords(colIndex, rowIndex, false));
                        }
                    }
                    return result;
                },
                // setValue
                // for LOV type of columns, use either composite value {v:??, d:??} or the v value
                setValue: function(rowIndex, columnName, value, suppressChangeEvent)
                {
                    let svColumn = columnsByName[columnName];
                    if (svColumn)
                    {
                        if (svColumn.lib4x.valueToDataValue)
                        {
                            value = svColumn.lib4x.valueToDataValue(value);
                        }                     
                        let colIndex = svColumn.index;
                        if (colIndex >= 0)
                        {
                            if (suppressChangeEvent === true)
                            {
                                instance.options.lib4x.suppressChangeEvent = true;
                            }
                            instance.setValueFromCoords(colIndex, rowIndex, value, true);
                            if (suppressChangeEvent === true)
                            {
                                instance.options.lib4x.suppressChangeEvent = false;
                            }                        
                        }
                    }
                },
                displayValueFor: function(columnName, value)
                {
                    let displayValue = '';
                    if (value != null && value !== '' && !Number.isNaN(value))
                    {
                        let svColumn = columnsByName[columnName];
                        if (svColumn)
                        {
                            if (svColumn.lib4x.displayValueFor)
                            {
                                displayValue = svColumn.lib4x.displayValueFor(value);
                            }  
                            else 
                            {
                                displayValue = String(value);
                            }  
                        }
                    }                 
                    return displayValue;
                },
                setReadOnly: function(rowIndex, columnName, isReadOnly)
                {
                    let colIndex = columnsByName[columnName]?.index;
                    if (colIndex >= 0)
                    {
                        let cellName = jspreadsheet.helpers.getCellNameFromCoords(colIndex, rowIndex);
                        if (cellName)
                        {
                            instance.setReadOnly(cellName, isReadOnly);
                        }
                    }
                },      
                isReadOnly: function(rowIndex, columnName)
                {
                    let colIndex = columnsByName[columnName]?.index;
                    return colIndex >= 0 ? instance.isReadOnly(colIndex, rowIndex) : undefined;
                },
                /*getRowData: function(rowIndex)
                {
                    let rowData = instance.getData()[rowIndex];
                    return rowData ? structuredClone(rowData) : null;
                },
                getRowMetadata: function(rowIndex)
                {
                    let rowMetadata = dsMetaUtil.getRowMeta(instance, rowIndex);  
                    return rowMetadata ? structuredClone (rowMetadata) : null; 
                },*/
                getColumns: function()
                {
                    let columns = [];
                    instance.options.columns.forEach((svColumn) => {
                        let column = {
                            name: svColumn.name,
                            label: svColumn.title,
                            index: svColumn.index,
                            type: typeof svColumn.type === 'string' ? svColumn.type : svColumn.lib4x.columnType.replace('lib4x_', ''),
                            readOnly: svColumn.readOnly
                        }
                        if (svColumn.customFormat)
                        {
                            column.formatMask = svColumn.customFormat;
                        }
                        columns[svColumn.name] = column;
                    });
                    return columns;
                }
            };
            return worksheetInterface;            
        }

        function getCtxPrototype(instance)
        {
            let worksheet = getWorksheetInterface(instance);
            let ctxPrototype = 
            {
                getValue: function(columnName)
                {
                    columnName ??= this.columnName;  
                    return worksheet.getValue(this.rowIndex, columnName);
                },  
                setValue: function(columnName, value, suppressChangeEvent)    
                {
                    worksheet.setValue(this.rowIndex, columnName, value, suppressChangeEvent);
                },
                displayValueFor: function(columnName, value)
                {
                    return worksheet.displayValueFor(columnName, value);
                },
                setReadOnly: function(columnName, isReadOnly)
                {
                    worksheet.setReadOnly(this.rowIndex, columnName, isReadOnly);
                },
                isReadOnly: function(columnName)
                {
                    columnName ??= this.columnName;
                    return worksheet.isReadOnly(this.rowIndex, columnName);
                },
                /*getRowData: function()
                {
                    return worksheet.getRowData(this.rowIndex);
                },
                getRowMetadata: function()
                {
                    return worksheet.getRowMetadata(this.rowIndex);
                },*/
                getColumns: function()
                {
                    return worksheet.getColumns();
                }
            }
            return ctxPrototype;
        }

        function toCanonicalValue(svColumn, dataValue)
        {
            let canonicalValue = dataValue;
            if (svColumn.lib4x.dataToNativeValue)
            {
                canonicalValue = svColumn.lib4x.dataToNativeValue(dataValue);
            }
            else 
            {
                canonicalValue = util.getScalarValue(dataValue);
            }
            return canonicalValue;
        }

        /*
         * syncToModel
         * synchronize all SV changes to the IG model: inserts, updates and deletes
         * on synchronizing a SV row, the 'onSynchronizeRow' handler is called
         */
        let syncToModel = function (svStaticId) {       
            let svStaticIdSv = svStaticId + SV_EXT;  
            let igStaticId = sv_igStaticId[svStaticIdSv];    
            let ig$ = $('#' + igStaticId);
            let instance = getJWorksheet(svStaticIdSv);
            let config = svConfig[svStaticIdSv];                          
            function gotoGridRow(recordId)
            {
                if (recordId)
                {
                    apex.region(igStaticId).call('getViews').grid.view$.grid('gotoCell', recordId); 
                }                 
            }   
            function onBeforeSynchronizeRow()
            {
                // if at any point it proves needed, next event can be implemented,
                // which is aimed at enabling making changes to the SV row before the row gets synchronized
                //apex.event.trigger(ig$, 'lib4x_sv_before_synchronize_row');
            }   
            function onSynchronizeRow(model, recordId, record)
            {              
                let ctx = {
                    model: model,
                    recordId: recordId,
                    record: record
                } 
                fireSVEvent(igStaticId, 'onSynchronizeRow', ctx);
                // do validations for required, maxlength, minimum, maximum, valid number
                // for GV, these are also resolved at the view layer and not centralized on the model layer, so we have to repeat them here
                // though some of them are validated server-side also 
                // validation on min/max date skipped for now as lower priority and anyhow server-side validated
                // also maxlength is effectively only validated for text columns for now
                // also no check on pattern
                let recMetadata = model.getRecordMetadata(recordId);                  
                let fields = model.getOption('fields');
                if (fields && recMetadata)
                {
                    let columnsByName = instance.options.lib4x.columnsByName;
                    for (const [fieldName, field] of Object.entries(fields))
                    {  
                        if (modelUtil.isRecordField(model, field) && columnsByName.hasOwnProperty(field.property))
                        {   
                            let svColumn = instance.options.lib4x.columnsByName[field.property];
                            let maxlength = svColumn.lib4x.maxlength;
                            let minimum = svColumn.lib4x.minimum;
                            let maximum = svColumn.lib4x.maximum;
                            if (field.isRequired || maxlength != null || minimum != null || maximum != null || svColumn.lib4x.columnType == 'lib4x_number' || svColumn.lib4x.columnType == 'lib4x_date_picker')
                            {
                                let fieldMetadata = modelUtil.getFieldMetadata(recMetadata, fieldName, false);
                                if ((recMetadata.inserted || fieldMetadata?.changed) && !(fieldMetadata?.error))
                                {
                                    let value = model.getValue(record, fieldName);
                                    value = util.getDisplayValue(value);
                                    let requiredError = (field.isRequired && (value === '' || value == null));
                                    let maxlengthError = (maxlength && value?.length > maxlength);
                                    let valid = !(requiredError || maxlengthError);
                                    let numberError = false;
                                    let numberMessage = null;
                                    let dateError = false;
                                    let dateMessage = null;                                    
                                    if (valid && value != null)
                                    {
                                        if (svColumn.lib4x.columnType == 'lib4x_number')
                                        {
                                            value = svColumn.lib4x.toDataValue(value);
                                            if (isNaN(value))
                                            {
                                                numberError = true;
                                                valid = false;
                                                numberMessage = apex.lang.getMessage('APEX.NUMBER_FIELD.VALUE_INVALID');                                                
                                            }
                                            if (valid && (minimum != null || maximum != null))
                                            {
                                                if ( minimum != null && maximum != null ) {
                                                    if ( value < minimum || value > maximum ) {
                                                        numberError = true;
                                                        valid = false;
                                                        numberMessage = apex.lang.formatMessage('APEX.NUMBER_FIELD.VALUE_NOT_BETWEEN_MIN_MAX', apex.locale.formatNumber(minimum, svColumn.customFormat), apex.locale.formatNumber(maximum, svColumn.customFormat));
                                                    }
                                                } else if ( minimum != null ) {
                                                    if ( value < minimum ) {
                                                        numberError = true;
                                                        valid = false;
                                                        numberMessage = apex.lang.formatMessage('APEX.NUMBER_FIELD.VALUE_LESS_MIN_VALUE', apex.locale.formatNumber(minimum, svColumn.customFormat));
                                                    }
                                                } else if ( maximum != null ) {
                                                    if ( value > maximum ) {
                                                        numberError = true;
                                                        valid = false;
                                                        numberMessage = apex.lang.formatMessage('APEX.NUMBER_FIELD.VALUE_GREATER_MAX_VALUE', apex.locale.formatNumber(maximum, svColumn.customFormat));
                                                    }
                                                }  
                                            }                                              
                                        }
                                        else if (svColumn.lib4x.columnType == 'lib4x_date_picker')
                                        {
                                            try
                                            {
                                                let parsedValue = apex.date.parse(value, svColumn.customFormat);
                                            }
                                            catch(e) 
                                            {
                                                dateError = true;
                                                valid = false;
                                                let validExample = ( svColumn.lib4x.apexItem.element.attr('valid-example'))
                                                    ? svColumn.lib4x.apexItem.element.attr('valid-example')
                                                    : apex.date.format( new Date(), svColumn.customFormat);
                                                dateMessage = apex.lang.formatMessageNoEscape('APEX.DATEPICKER.VALUE_INVALID', validExample);                                                
                                            }                                             
                                        }
                                    }
                                    if (requiredError || maxlengthError || numberError || dateError)
                                    {
                                        let apexItem = apex.item(field.elementId);
                                        let message = apexItem.element.attr('data-valid-message');
                                        if (requiredError)
                                        {
                                            message = message || apex.lang.getMessage('APEX.PAGE_ITEM_IS_REQUIRED');
                                        }
                                        else if (maxlengthError)
                                        {
                                            // '#LABEL# must have a length of max '
                                            message = message || getMessage('MAX_LENGTH_MSG') + maxlength; 
                                        }
                                        else if (numberError)
                                        {
                                            message = message || numberMessage;
                                        }
                                        else if (dateError)
                                        {
                                            message = message || dateMessage;
                                        }                                        
                                        message = apex.util.applyTemplate(message, {
                                            placeholders: {
                                                LABEL: field.label
                                            }
                                        });
                                        model.setValidity('error', recordId, fieldName, message);
                                    }
                                }
                            }
                        }
                    }
                }             
            }  
            let model = apex.region(igStaticId).call('getViews').grid.model;
            // cleanup any previous sync issues
            if (sv_syncIssues[svStaticIdSv]?.length > 0)
            {
                sv_syncIssues[svStaticIdSv].forEach((syncIssue) => {
                    if (syncIssue.recordId)
                    {
                        let recMetadata = model.getRecordMetadata(syncIssue.recordId);
                        if (recMetadata)
                        {
                            if (syncIssue.fieldName)
                            {
                                let fieldHasWarning = modelUtil.getFieldMetaPropertyValue(recMetadata, syncIssue.fieldName, 'warning');
                                if (fieldHasWarning)
                                {
                                    let message = modelUtil.getFieldMetaPropertyValue(recMetadata, syncIssue.fieldName, 'message');
                                    if (message == syncIssue.message)
                                    {
                                        model.setValidity('valid', syncIssue.recordId, syncIssue.fieldName, null);
                                    }
                                }
                            }
                            else
                            {
                                if (recMetadata.warning)
                                {
                                    if (recMetadata.message === syncIssue.message)
                                    {
                                        model.setValidity('valid', syncIssue.recordId, null, null);
                                    }
                                }
                            }
                        }
                    }
                });
            };
            let syncIssues = [];
            sv_syncIssues[svStaticIdSv] = syncIssues;
            if (instance) {
                let svInserts = 0;
                let svUpdates = 0;
                let svDeletes = 0;
                let modelInserts = 0;
                let modelUpdates = 0;
                let modelDeletes = 0;
                let svColumnsUpdOrder = getColumnsUpdateOrder(instance.options.columns);
                let prevRecord = null;
                let currentIds = new Set();
                let recordsInserted = false;
                let gotoRecordId = null;
                let svData = instance.getData();
                svData.forEach((svRow, y) => {
                    let recordId = dsMetaUtil.getRecordId(instance, y);
                    let rowMeta = instance.options.lib4x.dsMeta.get(recordId);   
                    let svInserted = false;   
                    let fireSynchronizeEvents = (rowMeta.sv.inserted || rowMeta.sv.updated);
                    if (fireSynchronizeEvents)
                    {
                        onBeforeSynchronizeRow();
                    }
                    if (rowMeta.sv.inserted)
                    {
                        svInserts = svInserts + 1;
                        // insert after the previous processed record
                        if (model.allowAdd())
                        {
                            let modelRecordId = model.insertNewRecord(null, prevRecord, null);
                            svInserted = true;
                            modelInserts = modelInserts + 1;
                            gotoRecordId = gotoRecordId || modelRecordId;
                            recordsInserted = true;
                            dsMetaUtil.setRecordId(instance, y, modelRecordId);
                            instance.options.lib4x.dsMeta.delete(recordId);
                            instance.options.lib4x.dsMeta.set(modelRecordId, rowMeta);
                            recordId = modelRecordId;
                        }
                        else 
                        {
                            recordId = null;
                            // Row could not be added (not allowed)
                            let message = getMessage('ROW_NOT_ADDED');   
                            syncIssues.push({
                                recordId: null,
                                operation: OP_INSERT,
                                fieldName: null,
                                rowData: structuredClone(svRow),                                
                                message: message
                            });                             
                        }
                        delete rowMeta.sv.inserted;                        
                    }
                    // maintain current recordId's
                    // as a recordId in APEX can be composite (an array), use JSON.stringify
                    let record = null;
                    if (recordId)
                    {
                        currentIds.add(JSON.stringify(recordId));
                        record = model.getRecord(recordId);
                    }
                    let modelUpdate = false;
                    let svUpdated = false;
                    if (rowMeta.sv.updated)
                    {
                        if (record)
                        { 
                            svColumnsUpdOrder.forEach((x) => {
                                let svColumn = instance.options.columns[x];
                                if (dsMetaUtil.fieldChanged(rowMeta, ORIG_SV, svColumn.name))                   
                                {
                                    if (svColumn.lib4x && !svColumn.readOnly) {
                                        let fieldName = svColumn.name;
                                        let svValue = svRow[x];                                  
                                        // in case there was a value conversion while reading the data from the IG model, 
                                        // convert the value here in a reverse way
                                        if (svColumn.lib4x?.toModelValue)
                                        {
                                            svValue = svColumn.lib4x?.toModelValue(svValue);
                                        }
                                        let currentValue = model.getValue(record, fieldName);
                                        if (!util.equalValues(svValue, currentValue)) 
                                        {
                                            svUpdated = true;
                                            let recMetadata = model.getRecordMetadata(recordId);   
                                            let fieldAccess = modelUtil.recordFieldWritable(model, recMetadata, record, fieldName);                                         
                                            if (fieldAccess.writable)
                                            {
                                                model.setValidity('valid', recordId, fieldName);    // default to valid
                                                model.setValue(record, fieldName, Number.isNaN(svValue) ? 'NaN' : svValue);     // Number.isNaN: should not happen; just to be sure
                                                modelUpdate = true;
                                            }
                                            else
                                            {
                                                let message = `${svColumn.title} could not be modified to '${util.getDisplayValue(svValue)}' (${fieldAccess.reason})`;
                                                syncIssues.push({
                                                    recordId: recordId,
                                                    operation: OP_UPDATE,
                                                    fieldName: fieldName,
                                                    message: message
                                                });                                                 
                                            }
                                            gotoRecordId = gotoRecordId || recordId;
                                        }
                                    }
                                }
                            });   
                        }
                        delete rowMeta.sv.updated;                  
                    }
                    if (!svInserted)
                    {
                        if (svUpdated)
                        {
                            svUpdates = svUpdates + 1;
                        }
                        if (modelUpdate)
                        {
                            modelUpdates = modelUpdates + 1;
                        }
                    }
                    if (fireSynchronizeEvents && recordId && record)
                    {
                        onSynchronizeRow(model, recordId, record);
                    }
                    if ((svInserted || modelUpdate) && record)
                    {
                        apex.event.trigger('#'+igStaticId, 'lib4xendrecordedit', {model: model, record: record});
                    }
                    if (record)
                    {
                        prevRecord = record;
                    }
                });
                // delete any model records which are not there anymore in the worksheet
                let recordsToDelete = [];
                instance.options.lib4x.dsMeta.forEach((rowMeta, recordId, map) => {
                    if (rowMeta.sv.deleted)
                    {                                     
                        svDeletes = svDeletes + 1;
                        let record = model.getRecord(recordId);
                        if (record)
                        {
                            if (model.allowDelete(record))
                            {
                                recordsToDelete.push(record);
                                modelDeletes = modelDeletes + 1;
                            }
                            else
                            {
                                // Row could not be deleted (not allowed)
                                let message = getMessage('ROW_NOT_DELETED');
                                syncIssues.push({
                                    recordId: recordId,
                                    operation: OP_DELETE,
                                    fieldName: null,
                                    message: message
                                });   
                            }
                        } 
                        delete rowMeta.sv.deleted;
                    }              
                });
                model.deleteRecords(recordsToDelete);
                syncIssues.forEach(syncIssue => {
                    if (syncIssue.recordId)
                    {
                        model.setValidity('warning', syncIssue.recordId, syncIssue.fieldName, syncIssue.message);
                    }
                });                
                if (recordsInserted)
                {
                    // for records which are added to a page unequal to the current page, 
                    // IG shows these records also in the current page. To get rid of it,
                    // we refresh the grid view
                    setTimeout(()=>{
                        apex.region(igStaticId).call('getViews').grid.view$.grid('refresh');  
                        gotoGridRow(gotoRecordId);                                                
                    }, 10);
                }
                else
                {
                    gotoGridRow(gotoRecordId); 
                }
                /*if (svInserts + svUpdates + svDeletes > 0)
                {
                    apex.message.showPageSuccess((svInserts > 0 ? modelInserts + ' row(s) out of ' + svInserts + ' added.<br> ' : '') + 
                                                (svUpdates > 0 ? modelUpdates + ' row(s) out of ' + svUpdates + ' modified.<br> ' : '') +
                                                (svDeletes > 0 ? modelDeletes + ' row(s) out of ' + svDeletes + ' deleted. ' : ''));
                    // make it an Info message
                    $('#APEX_SUCCESS_MESSAGE .t-Alert--success').addClass('t-Alert--info').removeClass('t-Alert--success');                                                    
                }*/
            }
        }

        // derive an epoch utc value given a date
        // as used for pasting to the clipboard (html part)
        // in Excel, dates are basically formatted epoch numbers 
        function jssDateToEpochUTC(date) {
            const epochUTC = Date.UTC(1899, 11, 30);

            const utcValue =
                Date.UTC(
                date.getFullYear(),
                date.getMonth(),
                date.getDate(),
                date.getHours(),
                date.getMinutes(),
                date.getSeconds(),
                date.getMilliseconds()
                );

            return (utcValue - epochUTC) / 86400000;
        }

        // get a next value for tempRecordSeq
        function getTempRecordId()
        {
            tempRecordSeq = tempRecordSeq + 1;
            return TEMP_ID_PREFIX + tempRecordSeq;
        }

        // check if any SV data has changed
        let dataHasChanged = function(svStaticId)
        {
            let svStaticIdSv = svStaticId + SV_EXT;
            let instance = getJWorksheet(svStaticIdSv);
            if (instance) 
            {            
                for (const rowMeta of instance.options.lib4x.dsMeta.values()) 
                {                    
                    if ((rowMeta.sv.transient == null) && (rowMeta.sv.inserted || rowMeta.sv.updated || rowMeta.sv.deleted))
                    { 
                        return true;
                    }
                };               
            }
            return false;
        }

        // detect if there is an open cell editor and if so, close it
        // JSS might have an on blur evenhandler open which would also call closeEditor, however
        // upon calling closeEditor() this blur handler is removed so close won't happen twice
        function closeAnyEditor(instance)
        {
            function closeEditor(cell, colIndex)
            {                 
                instance.closeEditor(cell, true);
            }
            if (instance.edition != null)
            {            
                let selectedCells = instance.getSelected(false);
                if (selectedCells && selectedCells.length == 1)
                {
                    let cell = selectedCells[0].element;
                    if ($(cell).hasClass(C_EDITOR))
                    {
                        closeEditor(cell, selectedCells[0].x);
                    }
                }   
                else
                {
                    // if selection is gone, we can still check the lastSelection
                    let lastSelection = instance.options.lib4x.lastSelection;
                    if (lastSelection?.length == 4)
                    {
                        if (lastSelection[0] == lastSelection[2] && lastSelection[1] == lastSelection[3])
                        {
                            let cellName = jspreadsheet.helpers.getCellNameFromCoords(lastSelection[0], lastSelection[1]);
                            let cell = instance.getCell(cellName);  
                            if ($(cell).hasClass(C_EDITOR))
                            {
                                closeEditor(cell, lastSelection[0]);
                            }                                               
                        }
                    }
                }    
            }     
        }

        // utility function which can be used from browser console to inspect the SV dataset metdata
        let logRowMetadata = function(svStaticId)
        {
            let svStaticIdSv = svStaticId + SV_EXT;
            let instance = getJWorksheet(svStaticIdSv);
            if (instance) 
            {            
                let svData = instance.getData();
                svData.forEach((svRow, rowIndex) => {
                    let recordId = dsMetaUtil.getRecordId(instance, rowIndex);
                    let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);
                    console.log(rowIndex+1, recordId, rowMeta.gv, rowMeta.sv);
                });
                let deletedCount = 0;
                instance.options.lib4x.dsMeta.forEach((rowMeta, recordId, map) => {
                    if (rowMeta.sv.deleted)
                    {
                        if (deletedCount == 0)
                        {
                            console.log('Deleted rows:');
                        }
                        console.log(recordId, rowMeta.gv, rowMeta.sv);
                        deletedCount = deletedCount + 1;
                    }
                });
                if (deletedCount == 0)
                {
                    console.log('No rows deleted');
                }
                console.log('Ds Metadata:')
                instance.options.lib4x.dsMeta.forEach((rowMeta, recordId, map) => {
                    console.log(recordId, rowMeta.gv, rowMeta.sv);
                });
                console.log('Data has changed: ', dataHasChanged(svStaticId));
            } 
            else
            {
                console.log('No active JSS instance for SV: ' + svStaticId);
            }   
            return '-End-';        
        }

        // fetch all data (taking maxRows into account) and recreate the SV
        let loadAll = function (svStaticId) 
        {
            let svStaticIdSv = svStaticId + SV_EXT;
            let instance = getJWorksheet(svStaticIdSv);
            if (instance) 
            {
                if (!sv_loadAllInProgress[svStaticIdSv])
                {
                    let spinner$ = apex.util.showSpinner($('#' + svStaticId));
                    let svOverlay$ = setDialogOverlay(svStaticId, true);
                    let numberRowsBefore = instance.getData().length;
                    fetchAll(svStaticId, function(status){
                        if (status.done)
                        {
                            createSpreadsheetView(svStaticId, true);
                            instance = getJWorksheet(svStaticIdSv);
                            if (instance.getData().length != numberRowsBefore)
                            {
                                // All loaded
                                apex.message.showPageSuccess(getMessage('ALL_LOADED'));
                            }  
                        } 
                        else if (status.error)
                        {
                            // Not able to load
                            apex.message.showPageSuccess(getMessage('NOT_ABLE_TO_LOAD'));
                            $('#APEX_SUCCESS_MESSAGE .t-Alert--success').addClass('t-Alert--danger').removeClass('t-Alert--success'); 
                        }       
                        spinner$.remove();    
                        svOverlay$.remove();                                    
                    });
                }
            }
        }

        let switchPaginationType = function (svStaticId)
        {
            let svStaticIdSv = svStaticId + SV_EXT;
            let config = svConfig[svStaticIdSv];
            let instance = getJWorksheet(svStaticIdSv);
            if (instance.options.pagination == null)
            {
                config.paginationType = 'PT_PAGE';
            }
            else
            {
                config.paginationType = 'PT_SCROLL';
            }
            let actionsContext = apex.actions.findContext('IGSpreadsheetView', $('#' + svStaticIdSv)[0]);            
            let filterAction = actionsContext.lookup('radiogroup-filter-rows');
            if (filterAction)
            {
                filterAction.filterChoice = FILTER_ALL;
            }            
            createSpreadsheetView(svStaticId, true);
        }
        
        // fetch all
        function fetchAll(svStaticId, callback)
        {
            let svStaticIdSv = svStaticId + SV_EXT;
            let igStaticId = sv_igStaticId[svStaticIdSv];   
            let config = svConfig[svStaticIdSv];                     
            sv_loadAllInProgress[svStaticIdSv] = true;                
            let model = apex.region(igStaticId).call('getViews').grid.model;
            modelUtil.fetchAll(model, config.options.maxRows, function(status) {
                if (status.done || status.error)
                {
                    sv_loadAllInProgress[svStaticIdSv] = false;                        
                    callback(status);
                }
            });            
        }

        function setDialogOverlay(svStaticId, transparent)
        {
            let css = {position: 'absolute'};
            if (transparent) {
                css['background-color'] = 'unset';
            }
            return $('<div class="apex_wait_overlay"></div>').css(css).prependTo($('#' + svStaticId).closest('.ui-dialog'));
        }

        // implementation for the filter-rows radiogroup buttons (toolbar)
        // as to enable filtering all rows/modified rows/rows with issues
        function setRowsVisibiliy(svStaticIdSv)
        {
            let actionsContext = apex.actions.findContext('IGSpreadsheetView', $('#' + svStaticIdSv)[0]);
            let filterChoice = actionsContext.get('radiogroup-filter-rows');
            let instance = getJWorksheet(svStaticIdSv);
            instance.rows.forEach((row, index) => {
                if ([FILTER_MODIFIED, FILTER_ISSUES].includes(filterChoice))
                {
                    let recordId = $(row.element).data('id');
                    let rowMeta = dsMetaUtil.getRecordMeta(instance, recordId);
                    if (((filterChoice == FILTER_MODIFIED) && (rowMeta.gv.inserted || rowMeta.gv.updated || rowMeta.sv.inserted || rowMeta.sv.updated)) ||
                        ((filterChoice == FILTER_ISSUES) && rowMeta.issues))
                    {
                        instance.showRow(row.y);
                    }
                    else
                    {
                        instance.hideRow(row.y);
                    }
                }
                else
                {
                    instance.showRow(row.y);
                }
            });
        }

        // upon synchronizing SV changes back to the model, any 'updateGridColumnAfter' configuration
        // is taken into account
        // this function derives the effective columns update order
        // it will check for circular dependency
        function getColumnsUpdateOrder(columns) {
            // build a lookup map
            const map = new Map(columns.map((col, idx) => [col.name, { col, idx }]));

            const visited = new Set();
            const temp = new Set();
            const result = [];

            function visit(name) {
                if (visited.has(name)) return;
                if (temp.has(name)) {
                    throw new Error("Circular dependency involving: " + name);
                }
                temp.add(name);
                const entry = map.get(name);
                if (entry) {
                    const col = entry.col;
                    if (col.lib4x.updateGridColumnAfter)
                    {
                        for (const dep of col.lib4x.updateGridColumnAfter) {
                            if (map.has(dep)) {
                                visit(dep);
                            }
                        }
                    }
                }
                temp.delete(name);
                visited.add(name);
                if (entry)
                {
                    result.push(entry.idx);
                }
            }
            for (const col of columns) {
                visit(col.name);
            }
            return result;
        }        

        // set/unset cell change indicator
        function updateCellChangeMarking(instance, colIndex, rowIndex, currentValue) 
        {
            let cellName = jspreadsheet.helpers.getCellNameFromCoords(colIndex, rowIndex);
            let svColumn = instance.options.columns[colIndex];
            let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);
            if (dsMetaUtil.fieldMetaHasProperty(rowMeta, ORIG_SV, svColumn.name, 'origValue'))
            {
                let origValue = rowMeta.sv.fields[svColumn.name].origValue;
                let changed = (!util.equalValues(currentValue, origValue));
                rowMeta.sv.fields[svColumn.name].changed = changed;
                let cell = instance.getCell(cellName);  
                setCellChangedClass(cell, ORIG_SV, changed);
            }
            updateRowUpdateMarking(instance, rowIndex);             
        }

        // set/unset row change background color
        function updateRowUpdateMarking(instance, rowIndex)
        {
            let updated = rowHasChanges(instance, rowIndex);
            let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);
            rowMeta.sv.updated = updated.sv;    
            setRowUpdatedClass(instance, rowIndex, rowMeta);
        }        

        function setCellChangedClass(cell, origin, haveChangedClass)
        {
            haveChangedClass ? $(cell).addClass('a-GV-cell is-changed-'+origin) : $(cell).removeClass('is-changed-'+origin);        // a-GV-cell intentionally added for proper styling of change indicator
        }

        function setCellHighlightClass(cell, highlight, haveHighlightClass, showHighlight)
        {
            let off = showHighlight ? '' : 'off_';
            haveHighlightClass ? $(cell).addClass('hlc_' + off + highlight) : $(cell).removeClass('hlc_' + off + highlight);
        }        

        function setCellErrorOrWarningClass(cell, message, messageType, haveClass)
        {
            let cell$ = $(cell);
            if (haveClass)
            {
                cell$.addClass('a-GV-cell is-' + messageType + '-gv');
                if (message)
                {
                    cell$.attr('title', ' ').tooltip({
                        content: message,
                        tooltipClass: 'a-GV-tooltip is-' + messageType,
                        items: '.is-' + messageType + '-gv:not(.is-changed-sv)'
                    });  
                }                 
            }
            else
            {
                cell$.removeClass('is-' + messageType + '-gv');
                if (cell$.tooltip('instance'))
                {
                    cell$.tooltip('destroy');
                }
            }
        }        

        function setRowUpdatedClass(instance, rowIndex, rowMeta)
        {
            let row$ = $(instance.rows[rowIndex].element);
            if (rowMeta.sv.updated)
            {
                if (!rowMeta.sv.inserted)
                {
                    row$.addClass('is-updated-sv');
                }
            }
            else
            {
                row$.removeClass('is-updated-sv');         
                if (rowMeta.gv.updated)
                {
                    if (!rowMeta.gv.inserted)
                    {
                        row$.addClass('is-updated-gv');
                    }
                }
                else
                {
                    row$.removeClass('is-updated-gv');            
                }                    
            }                  
        }

        function setRowInsertedClass(instance, rowIndex, rowMeta)
        {
            let row$ = $(instance.rows[rowIndex].element);
            if (rowMeta.sv.inserted)
            {
                row$.addClass('is-inserted-sv');
            }
            else
            {
                row$.removeClass('is-inserted-sv');
                rowMeta.gv.inserted ? row$.addClass('is-inserted-gv') : row$.removeClass('is-inserted-gv');               
            }                    
        }   
        
        function setRowHighlightClass(instance, rowIndex, rowMeta, showHighlight)
        {
            let off = showHighlight ? '' : 'off_';
            let row$ = $(instance.rows[rowIndex].element);
            row$.addClass('hlr_' + off + rowMeta.gv.highlight);                    
        }        

        function setRowHasIssuesClass(instance, rowIndex, rowMeta)
        {
            let row$ = $(instance.rows[rowIndex].element);
            row$.addClass(C_HAS_ISSUES);
            let message = '';
            rowMeta.issues.forEach(issue => {
                if (message)
                {
                    message = message + '<br/>';
                }
                message = message + issue.message;
            })
            row$.attr('title', ' ').find('.jss_row').tooltip({
                content: message,
                tooltipClass: "a-GV-tooltip is-warning"
            });             
        }

        function rowHasChanges(instance, rowIndex) 
        {
            function hasChangesForOrigin(rowMeta, origin)
            {
                for (const fieldName in rowMeta[origin].fields) 
                {
                    if (rowMeta[origin].fields[fieldName].changed)   
                    {
                        return true;
                    }   
                }
                return false;           
            }
            let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);
            return {
                gv: hasChangesForOrigin(rowMeta, ORIG_GV),
                sv: hasChangesForOrigin(rowMeta, ORIG_SV)
            }
        }         

        function updateRowInsertMarking(instance, rows)
        {
            rows.forEach((row, index) => {
                let rowIndex = row.row;
                let rowMeta = dsMetaUtil.getRowMeta(instance, rowIndex);
                rowMeta.sv.inserted = true;
                setRowInsertedClass(instance, rowIndex, rowMeta);             
            })
        }        

        // derive nested headers as per ig column groups
        // ungrouped columns are merged into empty header(s)
        function deriveNestedHeaders(igColumnGroups, svColumns) 
        {
            let result = null;
            // check if there is at least one grouped column
            let hasGroupedColumns = svColumns.some(
                c => c.lib4x && c.lib4x.columnGroup
            );
            if (hasGroupedColumns) 
            {
                let headers = [];
                let currentGroupId = undefined; // group id or null
                let colspan = 0;

                function flush() 
                {
                    if (colspan === 0) return;
                    if (currentGroupId) 
                    {
                        let grp = igColumnGroups[currentGroupId] || {};
                        headers.push({
                            title: grp.label ?? grp.heading ?? '',
                            align: { start: 'left', end: 'right' }[grp.headingAlignment] ?? 'center',
                            colspan
                        });
                    } 
                    else 
                    {
                        // ungrouped columns → empty header
                        headers.push({
                            colspan
                        });
                    }
                    colspan = 0;
                }

                for (const col of svColumns) 
                {
                    const groupId = col.lib4x && col.lib4x.columnGroup ? col.lib4x.columnGroup : null;
                    if (groupId !== currentGroupId) 
                    {
                        flush();
                        currentGroupId = groupId;
                    }
                    colspan++;
                }
                flush();
                result = [headers];
            }
            return result; 
        }

        // some column types maintain a LovMap as to map a display value to a return value
        function deriveLovMapFromItem(svColumn)
        {
            let apexItem = svColumn.lib4x.apexItem;
            let lovMap = new Map();
            let lovHasDuplicateDisplayValues = false;
            function handlePair(displayValue, returnValue) {
                let mappedValue = lovMap.get(displayValue);
                if ((mappedValue !== undefined) && (mappedValue !== returnValue)) 
                {
                    lovHasDuplicateDisplayValues = true;
                } 
                else 
                {
                    lovMap.set(displayValue, returnValue);
                }
            }       
            if (apexItem.item_type == 'SELECT')
            {
                apexItem.element.find('option').each(function() {
                    let option$ = $(this);
                    let returnValue = option$.attr('value');
                    let displayValue = util.normalizeDisplayValue(option$.text());
                    handlePair(displayValue, returnValue);
                });                
            }
            else if (apexItem.item_type == 'RADIO_GROUP')
            {
                apexItem.element.find('input[type=radio]').each(function () {
                    let radio$ = $(this);
                    let returnValue = radio$.val();
                    let displayValue = util.normalizeDisplayValue(radio$.data('display') || radio$.attr('aria-label'));
                    handlePair(displayValue, returnValue);              
                });
            }
            svColumn.lib4x.lovMap = lovMap;            
            svColumn.lib4x.lovHasDuplicateDisplayValues = lovHasDuplicateDisplayValues;
            if (lovHasDuplicateDisplayValues)
            {
                // duplicate display values we can't handle as we won't be able to
                // determine unambiguously the return value                
                svColumn.readOnly = true;
            }
        }

        // for IG column of type radio group (> 2 columns), a select list 
        // is used in SV
        function deriveSelectFromRadioGroup(apexItem)
        {
            let radioGroup$ = apexItem.element;
            let firstRadio$ = radioGroup$.find('input[type=radio]').first();
            let select$ = $('<select>', {
                id: radioGroup$.attr('id'),
                name: firstRadio$.attr('name'),
                class: 'selectlist apex-item-select js-ignoreChange js-tabbable',
                'data-native-menu': 'false',
                size: 1,
                tabindex: -1
            });
            radioGroup$.find('input[type=radio]').each(function () {
                let radio$ = $(this);
                select$.append(
                    $('<option>', {
                        value: radio$.val(),
                        text: radio$.data('display') || radio$.attr('aria-label'),
                        selected: radio$.is(':checked')
                    })
                );
            });
            return select$;
        }

        function showSyncResultMessage(igStaticId)
        {
            let model = apex.region(igStaticId).call('getViews').grid.model;
            if (model.hasErrors() || modelUtil.modelHasWarnings(model))
            {
                // Changes Synchronized. Grid has errors/warnings
                apex.message.showPageSuccess(getMessage('SYNCHRONIZED_WITH_MESSAGES'));
                $('#APEX_SUCCESS_MESSAGE .t-Alert--success').addClass('t-Alert--warning').removeClass('t-Alert--success');   
            }
            else
            {
                // Changes Synchronized
                apex.message.showPageSuccess(getMessage('CHANGES_SYNCHRONIZED'));
            }            
        }

        function getNextVisibleRow(instance, x, y)
        {
            const obj = instance;
            x = parseInt(x);
            y = parseInt(y);
            let yOrig = y;
            for (let j = y + 1; j < obj.rows.length; j++) {
                if (obj.records[j][x].element.style.display != 'none' && obj.rows[j].element.style.display != 'none') {
                    if (obj.records[j][x].element.getAttribute('data-merged')) {
                        if (obj.records[j][x].element == obj.records[y][x].element) {
                            continue;
                        }
                    }
                    y = j;
                    break;
                }
            }
            if (instance.options.pagination && instance.whichPage(y) != instance.pageNumber)
            {
                y = yOrig;
            }
            return y;            
        }

        function getPreviousVisibleRow(instance, x, y)
        {
            const obj = instance;
            x = parseInt(x);
            y = parseInt(y);
            let yOrig = y;
            for (let j = y - 1; j >= 0; j--) {
                if (obj.records[j][x].element.style.display != 'none' && obj.rows[j].element.style.display != 'none') {
                    if (obj.records[j][x].element.getAttribute('data-merged')) {
                        if (obj.records[j][x].element == obj.records[y][x].element) {
                            continue;
                        }
                    }
                    y = j;
                    break;
                }
            }
            if (instance.options.pagination && instance.whichPage(y) != instance.pageNumber)
            {
                y = yOrig;
            }            
            return y;            
        }

        function getJSpreadsheet(svStaticIdSv)
        {
            return $('#' + svStaticIdSv).data('jspreadsheet');
        }

        function getJWorksheet(svStaticIdSv)
        {
            return getJSpreadsheet(svStaticIdSv)?.worksheets[0];
        }

        function destroySpreadsheet(svStaticId)
        {
            let svStaticIdSv = svStaticId + SV_EXT;
            let actionsContext = apex.actions.findContext('IGSpreadsheetView', $('#' + svStaticIdSv)[0]);            
            let instance = getJWorksheet(svStaticIdSv);
            if (instance)
            {
                // remove reference to date picker / select one editiors in apex.items
                let svColumns = instance.options.columns;
                svColumns.forEach((svColumn, colIndex) => {
                    if ((svColumn.lib4x?.columnType == 'lib4x_date_picker') || (svColumn.lib4x?.columnType == 'lib4x_lov_one'))
                    {
                        let editor$ = svColumn.lib4x.editor$;
                        if (editor$)
                        {
                            let component$ = editor$.find('a-date-picker, a-select');
                            if (component$.length)
                            {
                                let id = component$.attr('id');
                                delete apex.items[id];
                                // component$ itself will get removed upon jss destroy
                            }
                        }
                    }
                });
                // store paginationPageSize
                if (localStorage) 
                {
                    let pageSize = null;
                    if ($('#' + svStaticIdSv + ' .jss_pagination_dropdown:visible').length > 0)
                    {
                        pageSize = $('.jss_pagination_dropdown').val();
                    }                  
                    if (Number(pageSize) > 0)
                    {
                        localStorage.setItem('paginationPageSize', pageSize);
                    }                   
                }                
            }     
            // the below will also remove C_LIB4X_SV_EDITORS_CONTAINER
            jspreadsheet.destroy($('#' + svStaticIdSv)[0]);  
            $('#' + svStaticIdSv).removeData('jspreadsheet');                          
        }

        return {
            initSV: initSV,
            syncToModel: syncToModel,
            openSpreadsheetView: openSpreadsheetView,
            getWorksheetInterface: getWorksheetInterface,
            logRowMetadata: logRowMetadata
        }
    })();

    let modelUtil = {
        sparselyLoaded: function(model) {
            return (Object.keys(model._data).length != model._data.length);
        },
        isRecordField: function(model, modelField)
        {
            return (modelField.hasOwnProperty('index') && modelField.property != model.getOption('metaField')); 
        },             
        recordFieldsHaveMessages: function (recMetadata, messageType) {
            let hasMessages = false;
            if (recMetadata) {
                let fields = recMetadata.fields;
                if (fields) {
                    for (const field in fields) {
                        if (fields[field][messageType]) {
                            hasMessages = true;
                            break;
                        }
                    }
                }
            }
            return hasMessages;
        },
        recordHasMessages: function(model, recordId, messageType) {
            let recMetaData = model.getRecordMetadata(recordId);
            return (recMetaData[messageType] || this.recordFieldsHaveMessages(recMetaData, messageType));
        },  
        // APEX model API has a hasErrors method but no hasWarnings or hasMessages method
        modelHasWarnings: function(model)  
        {
            let hasWarnings = false;
            model.forEach(function (record, index, recordId) {
                if (!hasWarnings)
                {
                    let recMeta = model.getRecordMetadata(recordId);
                    if (!recMeta.deleted && !recMeta.agg)
                    {           
                        hasWarnings = modelUtil.recordHasMessages(model, recordId, 'warning');
                    }
                }
            })
            return hasWarnings;
        },      
        getFieldMetadata: function(recMetadata, fieldName, createIfNotExists) {
            let result = null;
            if (recMetadata) {
                let fields = recMetadata.fields || (createIfNotExists ? recMetadata.fields = {} : null);
                if (fields) {
                    result = fields[fieldName] || (createIfNotExists ? fields[fieldName] = {} : null);
                }
            }           
            return result;  
        },
        getFieldMetaPropertyValue: function(recMetadata, fieldName, metaProperty) {
            let result = null;
            let fieldMetadata = this.getFieldMetadata(recMetadata, fieldName, false);
            if (fieldMetadata) {
                result = fieldMetadata[metaProperty];
            }
            return result;
        },
        recordFieldWritable: function(model, recMetadata, record, fieldName)
        {
            let result = {writable: true};
            if (!model.allowEdit(record))
            {
                result.writable = false;
                result.reason = 'row not editable';
            }
            else if (this.getFieldMetaPropertyValue(recMetadata, fieldName, 'ck'))
            {
                result.writable = false;
                result.reason = 'Grid cell is read only';                    
            }
            else if (modelUtil.getFieldMetaPropertyValue(recMetadata, fieldName, 'disabled'))
            {
                result.writable = false;
                result.reason = 'Grid cell is disabled';                      
            }
            return result;        
        },
        // fetchAll: based on a copy of model.fetchAll, customized as to stop loading when maxRows has been reached
        fetchAll: function(model, maxRows, pCallback, pNoProgress = true) {
            const o = model._options,
                savePageSize = o.pageSize,
                scopeName = "m" + model.modelId();
            let count, loadingIndicator, loading$,
                offset = 0;

            function load() {
                let r = null;
                if (offset >= maxRows)      // added for SV
                {
                    r = false;
                }
                else
                {
                    r = model.fetch( offset, function( err ) {
                        if (err) {
                            pCallback({
                                error: err
                            });
                        } else {
                            pCallback({
                                offset: offset,
                                total: model.getTotalRecords(),
                                done: false
                            });
                            offset += count;
                            load();
                        }
                    }, true ); // no progress because want progress to bracket all the requests
                }
                if ( r === null ) {
                    // request in progress wait and try again
                    setTimeout( function() {
                        load();
                    }, 500 );
                } else if ( r === false ) {
                    // restore the page size to what it was
                    o.pageSize = savePageSize;
                    if ( loadingIndicator ) {
                        apex.util.delayLinger.finish( scopeName, function() {
                            if ( typeof loading$ === 'function' ) {
                                loading$();
                            } else {
                                loading$.remove();
                            }
                        } );
                    }
                    // done
                    pCallback({
                        offset: offset,
                        total: model.getTotalRecords(),
                        done: true
                    });
                }
            }

            /*
            * The pageSize is intended to get small amounts of data over time. Enough to satisfy view requests
            * balanced with the likelihood that the user may never bother to look at it all.
            * But in this case all the data is to be fetched at once so it makes sense to request more data
            * to reduce the number of requests.
            * This assumes that no other fetch request will interrupt these fetchAll initiated fetch requests
            */
            if ( savePageSize < 1000 ) {
                o.pageSize = 1000;
            }
            count = o.pageSize;

            if ( !pNoProgress ) {
                const loadingIndicatorTmpl$ = $( '<span class="u-Processing u-Processing--inline"><span class="u-Processing-spinner"></span></span>' );

                loadingIndicator = makeLoadingIndicatorFunction( model );
                if ( loadingIndicator ) {
                    apex.util.delayLinger.start( scopeName, function () {
                        loading$ = loadingIndicator( loadingIndicatorTmpl$ );
                    } );
                }
            }

            while ( model._data[offset] ) {
                offset += 1;
            }

            load();
        }            
    };

    let dsMetaUtil = {
        initRowMeta: function()
        {
            let rowMeta = {
                gv: {
                    fields:{}
                },
                sv: {
                    fields:{}
                }
            };  
            return rowMeta;          
        },
        getRecordId: function(instance, rowIndex)
        {
            return $(instance.rows[rowIndex].element).data('id');
        },
        setRecordId: function(instance, rowIndex, recordId)
        {
            $(instance.rows[rowIndex].element).data('id', recordId);
        },
        getRecordMeta: function(instance, recordId)
        {
            return instance.options.lib4x.dsMeta.get(recordId);
        },
        getRowMeta: function(instance, rowIndex)
        {
            let recordId = this.getRecordId(instance, rowIndex);
            return this.getRecordMeta(instance, recordId);
        },
        setFieldMeta: function(rowMeta, origin, fieldName, property, value)
        {
            rowMeta[origin].fields[fieldName] = rowMeta[origin].fields[fieldName] || {};
            rowMeta[origin].fields[fieldName][property] = value;
        },
        fieldMetaHasProperty: function(rowMeta, origin, fieldName, property)
        {
            return(rowMeta[origin].fields[fieldName]?.hasOwnProperty(property));
        },
        fieldChanged: function(rowMeta, origin, fieldName)
        {
            let result = false;
            if (origin)
            {
                result = rowMeta[origin].fields[fieldName]?.changed;
            }
            else
            {
                result = ((rowMeta.gv.fields[fieldName]?.changed) || (rowMeta.sv.fields[fieldName]?.changed));
            }
            return result;
        },
        addIssue: function(rowMeta, issue)
        {
            if (!rowMeta.issues)
            {
                rowMeta.issues = [];
            }
            rowMeta.issues.push(issue);      
        }
    }

    // ==util module
    let util = {
        item:
        {
            cloneWebComponent: function (itemId, suffix) {
                const src$ = apex.item(itemId).element;
                const src = src$[0];
                const tag = src.tagName.toLowerCase();
                let attrs = [];
                for (const attr of src.attributes) {
                    let name = attr.name;
                    let value = attr.value;
                    if (name === "id" || name === "name") {
                        value += suffix;
                    }
                    attrs.push(`${name}="${value}"`);
                }
                // build HTML string — the id already set before construction
                // from which the element gets correctly registered as an apex item
                const html = `<${tag} ${attrs.join(" ")}></${tag}>`;
                return $(html);
            }
        },
        ig:
        {
            getTitle: function(igStaticId)
            {
                let igTitle = $('#' + igStaticId).find('.t-IRR-title').text();
                if (!igTitle)
                {
                    igTitle = apex.region(igStaticId).call('option').config?.regionAccTitle;   
                }  
                return igTitle;
            }
        },
        getScalarValue: function(value)
        {
            if (value !== null && typeof value === "object" && value.hasOwnProperty( "v" ))
            {
                value = value.v;
                if (Array.isArray(value))
                {
                    // Future; currently SV is not supporting multiple values so it should not land up here
                    value = 'Multiple';
                }
            }
            return value;
        },        
        getDisplayValue: function (value) {
            if (value !== null && typeof value === "object" && value.hasOwnProperty("d")) {
                value = value.d;
                if (Array.isArray(value)) 
                {
                    // Future; currently SV is not supporting multiple values so it should not land up here
                    value = 'Multiple';
                }
            }
            return value;
        },
        normalizeDisplayValue: function(value) {
            if (value == null) {
                return null;
            }
            return value
                .toString()
                .trim()                 // remove leading/trailing whitespace
                .replace(/\s+/g, " ")   // collapse internal whitespace
        },
        equalValues: function (value1, value2)
        {
            value1 = this.getScalarValue(value1) ?? '';
            value2 = this.getScalarValue(value2) ?? '';
            return value1 === value2;
        },
        valueNotEmpty: function(value)
        {
            return value !== null && value !== undefined && value !== '';
        },
        valueIsString: function(value)
        {
            return (typeof value === 'string' || value instanceof String);
        },
        toArray: function(value) 
        {
            if (value == null) return [];
            return Array.isArray(value) ? value : [value];
        },
        resolveValue: function(value, ...args) {
            return (typeof value === 'function') ? value(...args) : value;
        },
        // oracleNumberMaskToJSSMask: currently not used; could be used for number formula columns
        oracleNumberMaskToJSSMask: function(oracleMask) 
        {
            function normalizeJSSMask(mask, groupSep) {
                let esc = groupSep.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                let match = mask.match(/^([^#0]*)(.*)$/);
                if (!match) return mask;
                let prefix = match[1];
                let numeric  = match[2];
                numeric = numeric.replace(new RegExp(`^#+${esc}`), `#${groupSep}`);
                return prefix + numeric;
            }
            let groupSep   = apex.locale.getGroupSeparator();
            let decimalSep = apex.locale.getDecimalSeparator();
            let currency   = apex.locale.getCurrency();
            let isoCurrency   = apex.locale.getISOCurrency();
            let mask = oracleMask.toUpperCase();
            // remove Oracle specific tokens:
            mask = mask.replace(/FM/g, '');
            mask = mask.replace(/PR|MI|S/g, '');
            mask = mask.replace(/V\d+/g, '');            
            // currency (L or C → symbol)
            mask = mask.replace(/L/g, currency);
            mask = mask.replace(/C/g, isoCurrency);
            // group and decimal separators
            mask = mask.replace(/G/g, groupSep);
            mask = mask.replace(/D/g, decimalSep);
            // digits
            mask = mask.replace(/9/g, '#');
            mask = mask.replace(/0/g, '0');
            // remove everything that is not meaningful for JSpreadsheet
            // (spaces, leftover Oracle directives)
            mask = normalizeJSSMask(mask.trim(), apex.locale.getGroupSeparator());
            return mask;
        },
        oracleNumberMaskToMsoNumberFormat: function(oracleMask) 
        {
            if (!oracleMask) return 'General';
            let fmt = oracleMask.toUpperCase();
            // remove FM (Excel has no padding anyway)
            fmt = fmt.replace(/FM/g, "");
            // detect negative styles
            const hasMI = /MI/.test(fmt);
            const hasPR = /PR/.test(fmt);
            // remove MI / PR tokens
            fmt = fmt.replace(/MI|PR/g, "");
            let currency = apex.locale.getCurrency();
            let isoCurrency = apex.locale.getISOCurrency();
            let dualCurrency = apex.locale.getDualCurrency();
            // Oracle → Excel basic replacements
            fmt = fmt
                .replace(/G/g, ",")                                             // group separator
                .replace(/D/g, ".")                                             // decimal separator
                .replace(/9/g, "#")                                             // optional digits 
                .replace(/L/g, currency ?? '')                                  // local currency symbol
                .replace(/C/g, isoCurrency?.replace(/./g, '\\\\$&') ?? '')      // iso currency symbol
                .replace(/U/g, dualCurrency ?? '');                             // dual currency symbol
            // clean multiple dots (Oracle allows weird combos)
            fmt = fmt.replace(/\.{2,}/g, ".");
            // trim
            fmt = fmt.trim();
            // handle negatives
            if (hasMI) 
            {
                // positive;negative-with-minus-at-end
                return `${fmt};${fmt}-`;
            }
            if (hasPR) 
            {
                // positive;negative-in-angle brackets
                // by using angle brackets, we can also paste back to SV
                return `${fmt};<${fmt}>`;
            }
            // default: Excel auto handles negatives
            return fmt;
        },
        lovMapValueForReturnValue: function(lovMap, value)
        {
            if (value && util.valueIsString(value))
            {
                let entry = [...lovMap.entries()].find(([k, v]) => v === value);  
                if (entry)
                {
                    value = {v: value, d: entry[0]};
                }                          
                else
                {
                    value = {v: '', d: ''};
                }
            } 
            return value;           
        },
        // check if event is a keypress of a regular character, so excluding function keys and other non-printable keys
        isPrintableKey: function(event) {
            return (
                event instanceof KeyboardEvent &&
                event.key.length === 1 &&
                !event.ctrlKey &&
                !event.metaKey &&
                !event.altKey
            );
        }        
    };

    function oracleDateMaskToMsoNumberFormat(oracleFormat) 
    {
        if (/\bDS\b/g.test(oracleFormat)) 
        {
            oracleFormat = oracleFormat.replace(/\bDS\b/ig, apex.locale.getDSDateFormat());
        }
        if (/\bDL\b/g.test(oracleFormat)) 
        {
            oracleFormat = oracleFormat.replace(/\bDL\b/ig, apex.locale.getDLDateFormat());
        }
        const tokens = tokenizeOracleFormat(oracleFormat);
        const excelFmt = mapTokensToExcel(tokens);
        let locale = apex.locale.getLanguage();
        return locale ? `[$-${locale}]${excelFmt}` : excelFmt;
    }

    // tokenizer: splits format into tokens and literals, supports FM modifier 
    function tokenizeOracleFormat(fmt) 
    {
        const tokens = [
            'YYYY','RRRR','RR','YY',
            'MONTH','MON','MM',
            'DAY','DY',
            'DD',
            'HH24','HH12','HH',
            'MI','SS',
            'AM','PM',
            'DS','DL'
        ];
        // sort longest first
        tokens.sort((a,b) => b.length - a.length);
        let result = [];
        let i = 0;
        let fmActive = false;
        const upper = fmt.toUpperCase();
        while (i < fmt.length) 
        {
            // FM modifier (case-insensitive)
            if (upper.substr(i,2) === 'FM') 
            {
                fmActive = true;
                i += 2;
                continue;
            }
            // quoted literal
            if (fmt[i] === '"') 
            {
                let end = fmt.indexOf('"', i+1);
                if (end === -1) end = fmt.length - 1;
                result.push({
                    type: 'literal',
                    value: fmt.substring(i, end+1)
                });
                i = end + 1;
                continue;
            }
            // token match
            let matched = false;
            for (const token of tokens) 
            {
                if (upper.startsWith(token, i)) 
                {
                    result.push({
                        type: 'token',
                        value: token,
                        fm: fmActive
                    });
                    fmActive = false; // reset FM after token
                    i += token.length;
                    matched = true;
                    break;
                }
            }
            if (matched) continue;
            // single-char literal
            result.push({
                type: 'literal',
                value: fmt[i]
            });
            i++;
        }
        return result;
    }

    // mapper: converts tokens to Excel equivalents, respects FM rules 
    function mapTokensToExcel(tokens) 
    {
        // Oracle → Excel mapping
        const map = {
            YYYY: 'yyyy',
            RRRR: 'yyyy',
            RR: 'yy',
            YY: 'yy',

            MONTH: 'mmmm',
            MON: 'mmm',
            MM: 'mm',

            DAY: 'dddd',
            DY: 'ddd',

            DD: 'dd',

            HH24: 'hh',
            HH12: 'hh',
            HH: 'hh',

            MI: 'mm',
            SS: 'ss',

            AM: 'AM/PM',
            PM: 'AM/PM'
        };
        let result = '';
        for (const t of tokens)
        {
            if (t.type === 'literal') 
            {
                result += t.value;
                continue;
            }
            // map token
            if (map[t.value]) 
            {
                // FM is only relevant for text tokens (DAY, DY, MONTH, MON)
                if (t.fm && ['DAY','DY','MONTH','MON'].includes(t.value)) 
                {
                    result += map[t.value]; // Excel ignores trailing blanks anyway
                } 
                else 
                {
                    result += map[t.value]; // numeric tokens: ignore FM
                }
            }
            // unsupported tokens ignored
        }
        return result;
    }

    function initMessages() 
    {
        // here we have the labels and messages for which the developer should be 
        // able to config translations in APEX
        apex.lang.addMessages({
            'LIB4X.IG.SV.SPREADSHEET': 'Spreadsheet',
            'LIB4X.IG.SV.SHEET': 'Sheet',
            'LIB4X.IG.SV.REVERT_ALL': 'Revert All',
            'LIB4X.IG.SV.REVERT_ALL_CHANGES': 'Revert All Changes',
            'LIB4X.IG.SV.NO_ROWS_DELETED_NOT_ALL_ALLOWED': 'No rows deleted - not all selected rows are allowed to be deleted',
            'LIB4X.IG.SV.DELETE_LAST_ROW_NOT_POSSIBLE': 'It is not possible to delete the last row',
            'LIB4X.IG.SV.DIALOG.MAXIMIZE': 'Maximize',
            'LIB4X.IG.SV.DIALOG.RESTORE': 'Restore',
            'LIB4X.IG.SV.DIALOG.OK': 'OK',
            'LIB4X.IG.SV.DIALOG.CANCEL': 'Cancel',
            'LIB4X.IG.SV.DIALOG.OPEN_CHANGES_Q_CLOSE': 'New changes will be lost. Close the dialog?',
            'LIB4X.IG.SV.AGGREGATE.TOTAL': 'Total',
            'LIB4X.IG.SV.AGGREGATE.AVERAGE': 'Average',
            'LIB4X.IG.SV.AGGREGATE.MINIMUM': 'Minimum',
            'LIB4X.IG.SV.AGGREGATE.MAXIMUM': 'Maximum',
            'LIB4X.IG.SV.GRID_ROW_HAS_ERROR': 'Grid row has validation error(s)',
            'LIB4X.IG.SV.GRID_ROW_HAS_WARNING': 'Grid row has warning message(s)',
            'LIB4X.IG.SV.NOT_ABLE_TO_OPEN': 'Not able to open the Spreadsheet View. Please save and refresh the data and try again.',
            'LIB4X.IG.SV.UNDO_LAST_CHANGE': 'Undo the last change',
            'LIB4X.IG.SV.REDO_RECENT_CHANGE': 'Redo the most recent change',
            'LIB4X.IG.SV.ADD_BEFORE': 'Add Before',
            'LIB4X.IG.SV.ADD_ROW_BEFORE': 'Add a new row before',
            'LIB4X.IG.SV.ADD_AFTER': 'Add After',
            'LIB4X.IG.SV.ADD_ROW_AFTER': 'Add a new row after',
            'LIB4X.IG.SV.DELETE_ROWS': 'Delete Row(s)',
            'LIB4X.IG.SV.LOAD_ALL': 'Load All',
            'LIB4X.IG.SV.SYNCHRONIZE_FIRST': 'There are changes. Pls synchronize first.',
            'LIB4X.IG.SV.SYNCHRONIZE_WITH_GRID': 'Synchronize with Grid',
            'LIB4X.IG.SV.EDIT_ON_FOCUS': 'Edit on Focus',
            'LIB4X.IG.SV.SHOW_ALL': 'Show all',
            'LIB4X.IG.SV.SHOW_MODIFIED_ROWS': 'Show modified rows',
            'LIB4X.IG.SV.SHOW_ISSUES_ROWS': 'Show rows with issues',
            'LIB4X.IG.SV.SWITCH_PAGINATION': 'Switch Scroll/Page Pagination',
            'LIB4X.IG.SV.TOGGLE_HIGHLIGHTING': 'Toggle Highlighting',
            'LIB4X.IG.SV.HELP': 'Help',
            'LIB4X.IG.SV.HELP_TITLE': 'Interactive Grid Spreadsheet View',
            'LIB4X.IG.SV.ROW_NOT_ADDED': 'Row could not be added (not allowed)',
            'LIB4X.IG.SV.ROW_NOT_DELETED': 'Row could not be deleted (not allowed)',
            'LIB4X.IG.SV.MAX_LENGTH_MSG': '#LABEL# must have a length of max ',
            'LIB4X.IG.SV.ALL_LOADED': 'All loaded',
            'LIB4X.IG.SV.NOT_ABLE_TO_LOAD': 'Not able to load',
            'LIB4X.IG.SV.SYNCHRONIZED_WITH_MESSAGES': 'Changes Synchronized. Grid has errors/warnings',
            'LIB4X.IG.SV.CHANGES_SYNCHRONIZED': 'Changes Synchronized',
            'LIB4X.IG.SV.SEARCH': 'Search',
            'LIB4X.IG.SV.Q_DELETE_ROWS': 'Are you sure to delete the selected rows?'
        });
        jspreadsheet.setDictionary({
            'Search': getMessage('SEARCH'),
            'Are you sure to delete the selected rows?': getMessage('Q_DELETE_ROWS')
        });            
    }

    function getMessage(key) {
        return apex.lang.getMessage('LIB4X.IG.SV.' + key);
    }

    /*
     * Main plugin init function
     * enableFor: enable for all IG's, a specific IG, or IG's having filterClass
     */
    let init = function (svStaticId, enableFor, igStaticId, filterClass, paginationType, initFunc) 
    {
        if (enableFor == 'ENBL_ALL')
        {
            igStaticId = null;
            filterClass = null;
        }
        else if (enableFor == 'ENBL_SPECIFIC')
        {
            filterClass = null;
        }
        else if (enableFor == 'ENBL_CLASS')
        {
            igStaticId = null;
        }        
        initMessages();
        let svStaticIdSv = svStaticId + SV_EXT;
        sv_loadAllInProgress[svStaticIdSv] = false;
        sv_syncIssues[svStaticIdSv] = null;
        // default options
        let options = {
            maxRows: 5000,
            maxAdditionalRows: 1000,
            excludeCalculatedColumns: false,
            applyReadOnlyCells: false,
            applyHighlighting: true,
            rememberDialogCoordinates: false,
            helpText: null,
            additionalHelpText: null,
            buttons : {
                editOnFocus: true,
                ig: {
                    revertAll: false
                }
            }   
        };
        if (initFunc) {
            // call init function which the developer can use to 
            // programmatically specify the config option which are not available declaratively
            options = initFunc(options) || {};
        }
        let config = {};
        config.paginationType = paginationType;
        config.options = options;
        svConfig[svStaticIdSv] = config;
        // create region interface
        apex.region.create(svStaticId, {
            type: "IGSpreadsheetView",
            getWorksheet: function() {
                let instance = $('#' + svStaticIdSv).data('jspreadsheet')?.worksheets[0];
                return instance ? spreadsheetViewModule.getWorksheetInterface(instance) : null;
            }
        });
        spreadsheetViewModule.initSV(svStaticId);
        gridModule.initIGs(igStaticId, filterClass, svStaticId);
    };

    const [apexMajorVersion, apexMinorVersion, apexPatchVersion] = apex.env.APEX_VERSION.split('.').map(Number);
    // APEX 26.1 uses apex-core-font
    const iconFontFamily = apexMajorVersion >= 26 ? 'apex-core-font' : 'apex-5-icon-font';
    document.documentElement.style.setProperty('--lib4x-icon-font-family', iconFontFamily);

    window.lib4x = window.lib4x || {};
    lib4x.ig = lib4x.ig || {};
    lib4x.ig.spreadsheetView = lib4x.ig.sv || {};

    lib4x.ig.spreadsheetView.registerHandlers = function(igStaticId, eventHandlers){
        sv_eventHandlers[igStaticId] = eventHandlers;
    };
    lib4x.ig.spreadsheetView.unregisterHandlers = function(igStaticId){
        delete sv_eventHandlers[igStaticId];
    };

    return {
        _init: init,
        _logRowMetadata: spreadsheetViewModule.logRowMetadata
    }
})(apex.jQuery);
