#! python3
# # -*- coding: utf-8 -*-
# __title__ = "Arriero"                            # Name of the button displayed in Revit UI
# __doc__ = """Version = 0.5
# Date    = 2026-02-14
# Author: Aquelon - aquelon@pm.me 
# _____________________________________________________________________
# Description:
# DataExporterImporter for Rhino 8
# Exports and imports object User Keys/Values to/from Excel files using GUID tracking.
# _____________________________________________________________________
# How-to:
# -> Run the script in Rhino 8 (RunPythonScript)
# -> A window opens with Export and Import buttons and import options
# -> To export: Click "Export Data", select objects, choose save location
# -> To import: Configure options (backup, create keys, empty cells), click "Import Data"
# -> Select the Excel file, review the summary report
# -> Tip: Use "Apply only to pre-selected objects" to limit new key creation scope
# _____________________________________________________________________
# Last update:
# - [14.02.2026] - 0.5 RELEASE
# _____________________________________________________________________
# To-Do:
# - UI needs improvement, it is functional at the moment, but default look.
# - Default folder locations need to be updated to neutral locations (desktop or documents) instead of script folder.
# - Add function that erase keys on objects if they are not present in Excel file, this will allow to use the tool for cleaning up objects by erasing the columns in Excel.
# Update the logic of the import with empty cells, it seems tha is bit mixed right now. If the cells are emptied in an object, it is possible to erase the key in that object, but this logic may not be transparent for the user.
# _____________________________________________________________________
# r: openpyxl

import rhinoscriptsyntax as rs
import Rhino
import System
import datetime
import os
from System.Windows.Forms import (
    Form, Button, Label, CheckBox, TextBox, OpenFileDialog,
    DialogResult, FormBorderStyle, MessageBox, MessageBoxButtons,
    MessageBoxIcon, RadioButton, GroupBox, FolderBrowserDialog,
    SaveFileDialog
)
from System.Drawing import Point, Size, Font, FontStyle

# Try to import openpyxl for Excel operations
try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    print("Warning: openpyxl not available. Please install it using: pip install openpyxl")


class DataExporterImporterGUI(Form):
    """Main GUI for the Data Exporter/Importer tool"""
    
    def __init__(self):
        # Initialize base Form class first
        Form.__init__(self)

        self.Text = "Rhino Data Exporter/Importer"
        self.Size = Size(500, 420)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

        # Create UI elements
        self._create_controls()
        
    def _create_controls(self):
        """Create all UI controls"""
        y_pos = 20
        
        # Title
        title = Label()
        title.Text = "Select Operation:"
        title.Location = Point(20, y_pos)
        title.Size = Size(450, 20)
        title.Font = Font("Arial", 10, FontStyle.Bold)
        self.Controls.Add(title)
        y_pos += 30
        
        # Export button
        self.export_btn = Button()
        self.export_btn.Text = "1. Export Data (Select Objects → Excel)"
        self.export_btn.Location = Point(20, y_pos)
        self.export_btn.Size = Size(450, 40)
        self.export_btn.Click += self.on_export_click
        self.Controls.Add(self.export_btn)
        y_pos += 50
        
        # Import button
        self.import_btn = Button()
        self.import_btn.Text = "2. Import Data (Excel → Update Objects)"
        self.import_btn.Location = Point(20, y_pos)
        self.import_btn.Size = Size(450, 40)
        self.import_btn.Click += self.on_import_click
        self.Controls.Add(self.import_btn)
        y_pos += 60
        
        # Import Options GroupBox
        options_group = GroupBox()
        options_group.Text = "Import Options"
        options_group.Location = Point(20, y_pos)
        options_group.Size = Size(450, 220)
        self.Controls.Add(options_group)
        
        opt_y = 25
        
        # Backup checkbox
        self.backup_check = CheckBox()
        self.backup_check.Text = "Create backup before import"
        self.backup_check.Location = Point(10, opt_y)
        self.backup_check.Size = Size(400, 20)
        self.backup_check.Checked = True
        options_group.Controls.Add(self.backup_check)
        opt_y += 25
        
        # Create missing keys checkbox
        self.create_keys_check = CheckBox()
        self.create_keys_check.Text = "Create missing keys from Excel columns"
        self.create_keys_check.Location = Point(10, opt_y)
        self.create_keys_check.Size = Size(400, 20)
        self.create_keys_check.Checked = True
        self.create_keys_check.CheckedChanged += self.on_create_keys_changed
        options_group.Controls.Add(self.create_keys_check)
        opt_y += 25

        # Radio buttons for new key creation scope
        self.new_keys_all_radio = RadioButton()
        self.new_keys_all_radio.Text = "Apply to all objects in Excel file"
        self.new_keys_all_radio.Location = Point(30, opt_y)
        self.new_keys_all_radio.Size = Size(400, 20)
        self.new_keys_all_radio.Checked = True
        options_group.Controls.Add(self.new_keys_all_radio)
        opt_y += 22

        self.new_keys_selected_radio = RadioButton()
        self.new_keys_selected_radio.Text = "Apply only to pre-selected objects"
        self.new_keys_selected_radio.Location = Point(30, opt_y)
        self.new_keys_selected_radio.Size = Size(400, 20)
        options_group.Controls.Add(self.new_keys_selected_radio)
        opt_y += 30
        
        # Update empty cells checkbox
        self.update_empty_check = CheckBox()
        self.update_empty_check.Text = "Update values when Excel cell is empty"
        self.update_empty_check.Location = Point(10, opt_y)
        self.update_empty_check.Size = Size(400, 20)
        self.update_empty_check.Checked = False
        self.update_empty_check.CheckedChanged += self.on_update_empty_changed
        options_group.Controls.Add(self.update_empty_check)
        opt_y += 25
        
        # Placeholder label and textbox
        self.placeholder_label = Label()
        self.placeholder_label.Text = "Placeholder value (leave blank to DELETE the key):"
        self.placeholder_label.Location = Point(30, opt_y)
        self.placeholder_label.Size = Size(320, 20)
        self.placeholder_label.Enabled = False
        options_group.Controls.Add(self.placeholder_label)
        
        self.placeholder_text = TextBox()
        self.placeholder_text.Location = Point(350, opt_y)
        self.placeholder_text.Size = Size(80, 20)
        self.placeholder_text.Text = "-"
        self.placeholder_text.Enabled = False
        options_group.Controls.Add(self.placeholder_text)
        
    def on_create_keys_changed(self, sender, event):
        """Enable/disable radio buttons based on create keys checkbox state"""
        enabled = self.create_keys_check.Checked
        self.new_keys_all_radio.Enabled = enabled
        self.new_keys_selected_radio.Enabled = enabled

    def on_update_empty_changed(self, sender, event):
        """Enable/disable placeholder controls based on checkbox state"""
        enabled = self.update_empty_check.Checked
        self.placeholder_label.Enabled = enabled
        self.placeholder_text.Enabled = enabled
        
    def on_export_click(self, sender, event):
        """Handle export button click"""
        if not EXCEL_AVAILABLE:
            MessageBox.Show(
                "openpyxl library is not available.\nPlease install it using: pip install openpyxl",
                "Missing Dependency",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
            return
        
        self.DialogResult = DialogResult.OK
        self.Tag = "export"
        self.Close()
        
    def on_import_click(self, sender, event):
        """Handle import button click"""
        if not EXCEL_AVAILABLE:
            MessageBox.Show(
                "openpyxl library is not available.\nPlease install it using: pip install openpyxl",
                "Missing Dependency",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
            return
        
        self.DialogResult = DialogResult.OK
        self.Tag = "import"
        self.Close()


def export_data_to_excel():
    """
    Export selected objects' User Keys and Values to Excel file
    """
    # Prompt user to select objects
    objects = rs.GetObjects("Select objects to export data", preselect=True)
    if not objects:
        print("No objects selected. Export cancelled.")
        return
    
    print(f"Selected {len(objects)} objects for export.")
    
    # Collect all unique keys across all objects
    all_keys = set()
    object_data = []
    
    for obj in objects:
        guid = str(obj)
        keys = rs.GetUserText(obj)
        
        obj_dict = {"GUID": guid}
        
        if keys:
            for key in keys:
                value = rs.GetUserText(obj, key)
                obj_dict[key] = value if value else ""
                all_keys.add(key)
        
        object_data.append(obj_dict)
    
    # Sort keys alphabetically for consistent column order
    sorted_keys = sorted(all_keys)
    column_headers = ["GUID"] + sorted_keys
    
    print(f"Found {len(sorted_keys)} unique keys across all objects.")
    
    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Rhino Object Data"
    
    # Write headers
    for col_idx, header in enumerate(column_headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # Write data rows
    for row_idx, obj_dict in enumerate(object_data, start=2):
        for col_idx, header in enumerate(column_headers, start=1):
            value = obj_dict.get(header, "")
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Prompt user to select save location
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    default_filename = f"RhinoData_{timestamp}.xlsx"

    save_dialog = SaveFileDialog()
    save_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    save_dialog.FileName = default_filename
    save_dialog.Title = "Save Excel File"
    save_dialog.DefaultExt = "xlsx"

    if save_dialog.ShowDialog() != DialogResult.OK:
        print("Export cancelled by user.")
        return

    filepath = save_dialog.FileName

    # Save workbook
    try:
        wb.save(filepath)
        print(f"\nExport successful!")
        print(f"File saved: {filepath}")
        print(f"Exported {len(object_data)} objects with {len(sorted_keys)} keys.")

        MessageBox.Show(
            f"Export successful!\n\n"
            f"Objects: {len(object_data)}\n"
            f"Keys: {len(sorted_keys)}\n\n"
            f"File: {os.path.basename(filepath)}",
            "Export Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        MessageBox.Show(
            f"Error saving Excel file:\n{e}",
            "Export Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )


def import_data_from_excel(create_backup, create_missing_keys,
                           update_empty, placeholder, only_selected):
    """
    Import data from Excel file and update objects based on GUID

    Args:
        create_backup (bool): Create backup before import
        create_missing_keys (bool): Create keys that don't exist on objects
        update_empty (bool): Update values when Excel cell is empty
        placeholder (str): Placeholder value for empty cells (empty string means delete key)
        only_selected (bool): Only apply new keys to pre-selected objects
    """
    # Get pre-selected objects if only_selected is True
    selected_guids = None
    if only_selected and create_missing_keys:
        selected_objects = rs.SelectedObjects()
        if not selected_objects:
            MessageBox.Show(
                "No objects are selected.\n\n"
                "Please select objects first if you want to use\n"
                "'Apply only to pre-selected objects' option.",
                "No Selection",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return
        selected_guids = set(str(obj) for obj in selected_objects)
        print(f"Pre-selected {len(selected_guids)} objects for new key creation.")

    # Open file dialog to select Excel file
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Select Excel file to import"

    if dialog.ShowDialog() != DialogResult.OK:
        print("Import cancelled.")
        return

    filepath = dialog.FileName
    print(f"Selected file: {filepath}")
    
    # Create backup if requested
    if create_backup:
        result = MessageBox.Show(
            "Do you want to create a backup before importing?\n\n"
            "This will export all current object data to a new Excel file.",
            "Create Backup",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )

        if result == DialogResult.Yes:
            print("\nCreating backup before import...")
            # Get all objects in document for backup
            all_objects = rs.AllObjects()
            if all_objects:
                # Select all objects for backup
                rs.SelectObjects(all_objects)
                export_data_to_excel()
                rs.UnselectAllObjects()
                print("Backup created.")
            else:
                print("No objects found for backup.")
    
    # Load Excel file
    try:
        wb = load_workbook(filepath)
        ws = wb.active
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        MessageBox.Show(
            f"Error loading Excel file:\n{e}",
            "Import Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )
        return
    
    # Read headers (first row)
    headers = []
    for cell in ws[1]:
        if cell.value:
            headers.append(str(cell.value))
    
    if not headers or headers[0] != "GUID":
        print("Error: First column must be 'GUID'")
        MessageBox.Show(
            "Invalid Excel format.\nFirst column must be 'GUID'.",
            "Import Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )
        return
    
    # Keys are all columns except GUID
    keys = headers[1:]
    print(f"\nFound {len(keys)} keys in Excel file.")
    
    # Process each row
    updated_count = 0
    not_found_count = 0
    not_found_guids = []
    keys_created = 0
    keys_updated = 0
    keys_deleted = 0
    skipped_not_selected = 0
    
    for row_idx in range(2, ws.max_row + 1):
        guid_str = ws.cell(row=row_idx, column=1).value
        
        if not guid_str:
            continue
        
        # Try to find object by GUID
        try:
            guid = System.Guid(str(guid_str))
            obj = rs.coerceguid(guid)
            
            if not rs.IsObject(obj):
                not_found_count += 1
                not_found_guids.append(guid_str)
                continue

            # Check if this object can receive new keys
            can_create_keys = create_missing_keys
            object_skipped_selection = False
            if can_create_keys and selected_guids is not None:
                # Only allow key creation if object is in selected set
                if guid_str not in selected_guids:
                    can_create_keys = False
                    object_skipped_selection = True

            # Process each key/value pair
            for col_idx, key in enumerate(keys, start=2):
                excel_value = ws.cell(row=row_idx, column=col_idx).value
                
                # Check if value is None or empty string
                is_empty = excel_value is None or str(excel_value).strip() == ""
                
                if is_empty and not update_empty:
                    # Skip empty cells if update_empty is False
                    continue
                
                # Get current value on object
                current_value = rs.GetUserText(obj, key)
                key_exists = current_value is not None
                
                if is_empty and update_empty:
                    # Empty cell with update_empty enabled
                    if placeholder:
                        # Set to placeholder
                        if not key_exists and not can_create_keys:
                            # Skip if key doesn't exist and we can't create it
                            if object_skipped_selection:
                                skipped_not_selected += 1
                            continue
                        rs.SetUserText(obj, key, placeholder)
                        if key_exists:
                            keys_updated += 1
                        else:
                            keys_created += 1
                    else:
                        # Delete the key
                        if key_exists:
                            rs.SetUserText(obj, key, None)
                            keys_deleted += 1
                else:
                    # Non-empty value
                    value_str = str(excel_value)

                    if not key_exists and not can_create_keys:
                        # Skip if key doesn't exist and we can't create it
                        if object_skipped_selection:
                            skipped_not_selected += 1
                        continue

                    rs.SetUserText(obj, key, value_str)

                    if key_exists:
                        keys_updated += 1
                    else:
                        keys_created += 1
            
            updated_count += 1
            
        except Exception as e:
            print(f"Error processing GUID {guid_str}: {e}")
            not_found_count += 1
            not_found_guids.append(guid_str)
            continue
    
    # Show summary
    summary_msg = (
        f"Import Complete!\n\n"
        f"Objects updated: {updated_count}\n"
        f"Keys created: {keys_created}\n"
        f"Keys updated: {keys_updated}\n"
        f"Keys deleted: {keys_deleted}\n"
    )

    if skipped_not_selected > 0:
        summary_msg += f"Keys skipped (not in selection): {skipped_not_selected}\n"

    if not_found_count > 0:
        summary_msg += f"\nGUIDs not found: {not_found_count}\n"
        if len(not_found_guids) <= 5:
            summary_msg += "\n".join([f"  - {g}" for g in not_found_guids])
        else:
            summary_msg += "\n".join([f"  - {g}" for g in not_found_guids[:5]])
            summary_msg += f"\n  ... and {len(not_found_guids) - 5} more"
    
    print("\n" + "="*50)
    print(summary_msg)
    print("="*50)
    
    MessageBox.Show(
        summary_msg,
        "Import Summary",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information
    )


def main():
    """Main function to run the Data Exporter/Importer tool"""
    
    if not EXCEL_AVAILABLE:
        print("="*60)
        print("ERROR: openpyxl library is not installed")
        print("="*60)
        print("\nThis script requires the openpyxl library to work with Excel files.")
        print("Please install it using the following command:\n")
        print("  pip install openpyxl\n")
        print("If you're using Rhino 8's Python 3, you may need to:")
        print("  1. Open Windows Command Prompt (cmd)")
        print("  2. Navigate to Rhino's Python directory")
        print("  3. Run: python -m pip install openpyxl")
        print("="*60)
        return
    
    # Show GUI
    gui = DataExporterImporterGUI()
    result = gui.ShowDialog()
    
    if result == DialogResult.OK:
        operation = gui.Tag

        if operation == "export":
            export_data_to_excel()

        elif operation == "import":
            create_backup = gui.backup_check.Checked
            create_missing_keys = gui.create_keys_check.Checked
            update_empty = gui.update_empty_check.Checked
            placeholder = gui.placeholder_text.Text if update_empty else ""
            only_selected = gui.new_keys_selected_radio.Checked

            import_data_from_excel(
                create_backup,
                create_missing_keys,
                update_empty,
                placeholder,
                only_selected
            )


if __name__ == "__main__":
    main()
