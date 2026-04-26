# apex-ig-spreadsheet-view
Adds a Spreadsheet View to Interactive Grids for fast data editing, with support for copy-and-paste to and from Excel.

<p>
<img src="./ig-spreadsheetview.jpg" height="100%" width="100%">
</p>

### Configuration in Page Designer
To configure, create a new Region and select 'LIB4X - IG Spreadsheet View' as the Region Type. Select 'Inline Dialog' as the template. Select 'Dialogs, Drawers and Popups' as the Slot. On the Region Attributes, select for which IG/IG's you want to enable the Spreadsheet View.

<p>
<img alt="image" src="https://github.com/user-attachments/assets/8ce5f930-31e5-4f38-adac-0a8598fcfd04" height="40%" width="40%" />
</p>
<p>
In the IG toolbar, and extra item will show up, which starts the Spreadsheet View in a Modal Dialog.
</p>
<p>
<img alt="image" src="https://github.com/user-attachments/assets/8b80ff30-ae02-44b5-aed7-69418c75d30e" height="30%" width="30%" />
</p>

### Usage
The Spreadsheet View loads with a copy of the Grid data. You can edit it, add/delete rows, and copy/paste from and to Excel. Upon 'OK', the changes are synchronized back to the Grid. In the Grid, you can address any issues like validation errors and then save the data.

Upon selecting a cell, you can directly start typing to replace any current value. Or double-click or use F2 to change the existing value.

You can use the other familiar spreadsheet type of editing features like selecting cell(s), copy them and paste the values elsewhere. Or use the Fill Handle (bottom-right corner of a selection) to copy or fill into adjacent cells. You can also use copy-and-paste to/from Excel.

<ins>Edit on Focus</ins>: this button enables you to edit cells without need to first use F2 or double click.

<ins>Load All</ins>: initially, a subset of data might have been loaded only. Use this button to load all the data. It loads to a maximum of 5000 rows (configurable).

<ins>Synchronize</ins>: this button lets you synchronize your changes in between with the Grid without closing the dialog. Any resulting validation errors will be marked and shown in the spreadsheet.

Ctrl+Z/Ctrl+Y: shortcut keys for Undo/Redo. This will apply to changes which haven't been synchronized to the Grid yet.

### Technical Background
The plugin utilizes the [JSpreadsheet CE](https://bossanova.uk/jspreadsheet/) open-source spreadsheet component (JSS). JSS enables to add custom editors (custom column types). The IG Spreadsheet View plugin supports next column types:
- Text
- Number
- Date Picker
- Select List
- Checkbox
- Switch
- Radio group (2 options)
- Pill Buttons
- Select One

Here, Checkbox, Switch, Radio Group and Pill Buttons are so called Simple Choice columns, selecting out of 2 options. 
The Spreadsheet View inherits as much as possible from the IG column definitions. So Date Picker will look the same; a Number column will have same Format Mask (if any), etc. 
A Radio Group in IG will also be a Radio Group in Spreadsheet View if it has 2 options. Else, a Select List will be used. A Popup LOV column in IG will be a Select One column in Spreadsheet View. <br/>
Cascading List of Values are not supported - this is technically out of reach in a spreadsheet which has a much more free style of editing. Also Multi Value columns are not supported. Non supported IG columns are becoming Read-Only columns in Spreadsheet View.

IG Spreadsheet View loads the IG model data into an own copy of the data. It also maintains it's own metadata to keep track of what data has changed. Upon OK or Synchronize, the changes are updated back into the IG model data. This includes inserts and deletes. It will check if model updates are allowed. Also basic validations will be done, like Value Required, valid number, Min/Max number, valid date, etc. 

Here we come to an important point: in this whole process, no Column Item Dynamic Actions are executed! The model is the shared layer between the IG grid and the IG Spreadsheet View and updates do go via the model only. This can have implications. For example when you are using a DA to calculate a line total. You can resolve this by moving the calculation to the model layer and use the model [calcValue](https://docs.oracle.com/en/database/oracle/apex/24.2/aexjs/model.html#.FieldMeta) feature.

In general, for any (business) logic or any validations which you have implemented in the IG Grid UI layer, as to make that logic common between the IG Grid View and the IG Spreadsheet View, you have next options:
- move logic to the  model layer. A supporting plugin here is the IG Model Logic plugin, which makes implementing logic in the model layer much more convenient.
- make use of the 'onSynchronizeRow' event handler, which is fired by IG Spreadsheet View - see details below
- make use of the 'Execute Server-Side IG Row Logic' plugin

To the last option, an example is: 'Additional Columns'. In IG Grid, when you have a Popup LOV column, you might have additional columns populated. In IG Spreadsheet View, you don't have this type of interaction with servers-side data while editing. To still enable populating these additional columns, IG Spreadsheet View emits a 'Synchronize' event which you can select for a Dynamic Action. This event is fired after synchronization of changes to the model is complete. You can then use the DA here to 'Execute Server-Side IG Row Logic' as to read the values for the additional columns for all modified rows, and the plugin will populate the additional columns subsequently.

Needless to say, your regular server-side logic and validations offer the real protection of your data.

### Further Details
#### Options
In Attributes/Initialization JavaScript Function, you can configure some further options. For example the maximum number of rows which can be loaded:

```
function(options)
{
    options.maxRows = 7000;
    return options;
}
```

Next options can be configured:
- <ins>maxRows</ins> (number): the Spreadsheet View will open with the data as currently loaded into the IG model. By using the 'Load All' button, it will load all with a maximum of maxRows. The default is 5000.
- <ins>maxAdditionalRows</ins> (number): by using 'Add' button, new rows can be added to the sheet. Also new rows will be added in case you copy data from Excel where the number of rows exceeds the current end of the data in the IG Spreadsheet View. This option reflects the max number of rows which can be added. The default is 1000.
- <ins>excludeCalculatedColumns</ins> (boolean): default is false. If set to true, columns which have the [calcValue](https://docs.oracle.com/en/database/oracle/apex/24.2/aexjs/model.html#.FieldMeta) function configured will be skipped.
- <ins>applyReadOnlyCells</ins> (boolean): in case there are individual IG cells which are Read-Only (configured on model level via the fieldmetadata ck property), then by default these cells are not Read-Only in the IG Spreadsheet View. This fits to the optimistic editing mode of the spreadsheet view. So the default is false. You can override this behavior by setting this option to true.
- <ins>applyHighlighting</ins> (boolean): when the user has defined any Highlighting in the IG Grid, by default the Highlighting will also be shown in the Spreadsheet View. So default is true. You can switch it off with this option. There is also a button in the IG Spreadsheet View toolbar in case of Highlighting to enable the user to hide it.
- <ins>buttons.editOnFocus</ins> (boolean): default is true, but optionally you can get rid of this button using this setting
- <ins>rememberDialogCoordinates</ins> (boolean): by default, when opening the IG Spreadsheet View, it will use default coordinates (X, Y position, height, width). So default of this option is false. By setting it to yes, the dialog will open with previous coordinates.
- <ins>helpText</ins> (string): can be used in case you want to have your own help text behind the toolbar help button
- <ins>additionalHelpText</ins> (string): can be used to add extra help text in addition to the default help text
- <ins>buttons.ig.revertAll</ins> (boolean): default is false. Set to true in case you want a 'Revert All' button in the IG toolbar as to quickly undo all current changes (including inserts and deletes)

#### Event Handlers
You can register next event handlers: 'onChange' and 'onSynchronizeRow'. It's a programmatic construct. A context object is supplied with relevant details and some available methods on the prototype.<br/>
For 'onChange' (triggered upon cell value change):
<p>
<img width="40%" height="40%" alt="image" src="https://github.com/user-attachments/assets/a5c1f9d7-b6d6-4446-8573-3325bbcef2cb" />
</p>
For 'onSynchronizeRow':<br/>
<p>
<img width="80%" height="80%" alt="image" src="https://github.com/user-attachments/assets/ddb9fb2b-c6eb-40b6-aab3-162aa412c287" />   
</p>

```
$(function(){
    lib4x?.ig?.spreadsheetView?.registerHandlers('ig_order_lines', {
        onChange: function(ctx){
            if (ctx.columnName == 'QUANTITY' || ctx.columnName == 'UNIT_PRICE')  
            {          
                let quantity = ctx.getValue('QUANTITY');
                let unit_price = ctx.getValue('UNIT_PRICE');
                ctx.setValue('LINE_TOTAL', quantity * unit_price, true);                
            }         
        }
        /*,
        onSynchronizeRow: function(ctx)
        {

        }*/
    });
});
```

Notice, for 'onChange', native JS values are used for number/date method arguments.

#### DA Event
The 'Synchronize' event is emited when all changed rows are synchronized with the IG model. So you can utilize this event for any DA you want to execute:
<p>
<img width="40%" height="40%" alt="image" src="https://github.com/user-attachments/assets/047d86e5-4122-405b-9cae-1109c8c1a6ba" />
</p>

<h3>Plugin versions</h3>
Version 1.0.0 - build under APEX 24.2<br>

<h3>Third-Party Libraries and Licenses</h3>

This software uses:<br>
[JSpreadsheet CE](https://bossanova.uk/jspreadsheet/), which is licensed under the MIT License.<br>
Copyright (c) 2024 Jspreadsheet Ltd<br>
[License](https://github.com/jspreadsheet/ce/blob/master/LICENSE) | [Project](https://github.com/jspreadsheet/ce)
