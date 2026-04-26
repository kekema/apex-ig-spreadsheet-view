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

In general, for any (business) logic which you have implemented in the IG Grid UI layer, as to make that logic common between the IG Grid View and the IG Spreadsheet View, you have next options:
- move logic to the  model layer. A supporting plugin here is the IG Model Logic plugin, which makes implementing logic in the model layer much more convenient.
- make use of the 'onSynchronizeRow' event handler, which is fired by IG Spreadsheet View - see details below
- make use of the 'Execute Server-Side IG Row Logic' plugin

To the last option, an example is: 'Additional Columns'. In IG Grid, when you have a Popup LOV column, you might have additional columns populated. In IG Spreadsheet View, you don't have this type of interaction with servers-side data while editing. To still enable populating these additional columns, IG Spreadsheet View emits a 'Synchronize' event which you can select for a Dynamic Action. This event is fired after synchronization of changes to the model is complete. You can then use the DA here to 'Execute Server-Side IG Row Logic' as to read the values for the additional columns for all modified rows, and the plugin will populate the additional columns subsequently.








