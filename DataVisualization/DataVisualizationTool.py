"""
Data Visualization Tool - Unified Interface
A single GUI for the complete metadata visualization workflow.

FULLY SELF-CONTAINED - No external script dependencies

Author: RhinoGuire  
Version: 1.1
"""

import rhinoscriptsyntax as rs
import scriptcontext as sc
import csv
import os
from collections import defaultdict

# Import Eto forms
import Eto.Drawing as drawing
import Eto.Forms as forms


# ============================================================================
# SCRIPT 01 FUNCTIONS
# ============================================================================

def read_keys_from_csv(csv_path):
    """Read key definitions from CSV"""
    try:
        keys = {}
        with open(csv_path, 'r') as file:
            reader = csv.reader(file)
            next(reader, None)
            for row in reader:
                if len(row) >= 2:
                    key_name = row[0].strip()
                    default_value = row[1].strip() if row[1] else ''
                    if key_name:
                        keys[key_name] = default_value
        return keys
    except Exception as e:
        print("Error reading CSV: " + str(e))
        return None


def apply_keys_to_objects(objects, keys):
    """Apply keys to objects without overwriting"""
    stats = {'objects_processed': 0, 'keys_added': 0, 'keys_skipped': 0}
    for obj in objects:
        stats['objects_processed'] += 1
        for key_name, default_value in keys.items():
            if rs.GetUserText(obj, key_name) is None:
                rs.SetUserText(obj, key_name, default_value)
                stats['keys_added'] += 1
            else:
                stats['keys_skipped'] += 1
    return stats


# ============================================================================
# SCRIPT 02 FUNCTIONS
# ============================================================================

def collect_unique_values(objects):
    """Scan objects for unique key-value pairs"""
    key_values = defaultdict(set)
    for obj in objects:
        keys = rs.GetUserText(obj)
        if keys:
            for key in keys:
                value = rs.GetUserText(obj, key)
                if value is not None and value != '':
                    key_values[key].add(value)
    return {key: sorted(list(values)) for key, values in key_values.items()}


def write_color_map_csv(csv_path, key_value_dict):
    """Write color mapping CSV"""
    try:
        with open(csv_path, 'w') as file:
            writer = csv.writer(file)
            writer.writerow(['Key', 'Value', 'R', 'G', 'B', 'A'])
            rows = 0
            for key_name in sorted(key_value_dict.keys()):
                for value in key_value_dict[key_name]:
                    writer.writerow([key_name, value, '', '', '', ''])
                    rows += 1
        return rows
    except Exception as e:
        print("Error writing CSV: " + str(e))
        return 0


# ============================================================================
# SCRIPT 03 FUNCTIONS
# ============================================================================

def read_color_map_from_csv(csv_path):
    """Read color mappings from CSV"""
    color_map = {}
    try:
        with open(csv_path, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                key = row.get('Key', '').strip()
                value = row.get('Value', '').strip()
                r = row.get('R', '').strip()
                g = row.get('G', '').strip()
                b = row.get('B', '').strip()
                a = row.get('A', '').strip()
                if key and value:
                    try:
                        r_val = max(0, min(255, int(r) if r else 0))
                        g_val = max(0, min(255, int(g) if g else 0))
                        b_val = max(0, min(255, int(b) if b else 0))
                        a_val = max(0, min(255, int(a) if a else 255))
                        if key not in color_map:
                            color_map[key] = {}
                        color_map[key][value] = (r_val, g_val, b_val, a_val)
                    except ValueError:
                        continue
        return color_map
    except Exception as e:
        print("Error reading color CSV: " + str(e))
        return None


def analyze_objects_for_key(objects, active_key):
    """Analyze objects for unique values"""
    value_objects = {}
    objects_without_key = []
    for obj in objects:
        value = rs.GetUserText(obj, active_key)
        if value is None or value == '':
            objects_without_key.append(obj)
        else:
            if value not in value_objects:
                value_objects[value] = []
            value_objects[value].append(obj)
    return {'value_objects': value_objects, 'objects_without_key': objects_without_key}


def apply_colors_to_objects(objects, active_key, color_map, default_color):
    """Apply colors based on metadata"""
    stats = {
        'colored': 0,
        'default': 0,
        'missing_color_definition': [],
        'unused_color_definitions': []
    }
    
    key_color_map = color_map.get(active_key, {})
    analysis = analyze_objects_for_key(objects, active_key)
    value_objects = analysis['value_objects']
    objects_without_key = analysis['objects_without_key']
    used_values = set()
    
    for value, obj_list in value_objects.items():
        used_values.add(value)
        color = key_color_map.get(value)
        if color:
            for obj in obj_list:
                rs.ObjectColor(obj, (color[0], color[1], color[2]))
                stats['colored'] += 1
        else:
            for obj in obj_list:
                rs.ObjectColor(obj, default_color)
            stats['missing_color_definition'].append((value, obj_list))
    
    for obj in objects_without_key:
        rs.ObjectColor(obj, default_color)
        stats['default'] += 1
    
    csv_values = set(key_color_map.keys())
    unused = csv_values - used_values
    stats['unused_color_definitions'] = sorted(list(unused))
    
    return stats


# ============================================================================
# COLOR MANAGER DIALOG (SCRIPT 03)
# ============================================================================

class ColorManagerDialog(forms.Dialog[bool]):
    """Color Manager GUI"""
    
    def __init__(self, csv_path, objects):
        self.csv_path = csv_path
        self.objects = objects
        self.color_map = None
        self.active_key = None
        self.default_color = (128, 128, 128)
        self.problem_objects = []
        
        self.load_color_map()
        
        self.Title = "Color Manager"
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        self.Size = drawing.Size(500, 600)
        self.Topmost = True  # Keep on top
        
        self.create_controls()
        self.create_layout()
        
    def load_color_map(self):
        self.color_map = read_color_map_from_csv(self.csv_path)
        if not self.color_map:
            forms.MessageBox.Show("Error loading CSV", "Error")
            self.Close(False)
    
    def create_controls(self):
        self.key_label = forms.Label(Text="Select Metadata Key:")
        self.key_dropdown = forms.DropDown()
        
        available_keys = sorted(self.color_map.keys())
        for key in available_keys:
            self.key_dropdown.Items.Add(key)
        if available_keys:
            self.key_dropdown.SelectedIndex = 0
            self.active_key = available_keys[0]
        self.key_dropdown.SelectedIndexChanged += self.on_key_changed
        
        self.default_color_label = forms.Label(Text="Default Color:")
        self.default_color_picker = forms.ColorPicker()
        self.default_color_picker.Value = drawing.Color.FromArgb(128, 128, 128)
        self.default_color_picker.ValueChanged += self.on_default_color_changed
        
        self.legend_label = forms.Label(Text="Color Legend:")
        self.legend_panel = forms.Panel()
        self.legend_panel.Size = drawing.Size(450, 200)
        self.legend_panel.BackgroundColor = drawing.Colors.White
        self.update_legend()
        
        self.warnings_label = forms.Label(Text="Warnings & Errors:")
        self.warnings_text = forms.TextArea()
        self.warnings_text.ReadOnly = True
        self.warnings_text.Size = drawing.Size(450, 100)
        self.warnings_text.Text = "Click 'Update Colors' to analyze..."
        
        self.update_button = forms.Button(Text="Update Colors")
        self.update_button.Click += self.on_update_clicked
        
        self.select_button = forms.Button(Text="Select Problem Objects")
        self.select_button.Click += self.on_select_clicked
        self.select_button.Enabled = False
        
        self.close_button = forms.Button(Text="Close")
        self.close_button.Click += self.on_close_clicked
        
        self.status_label = forms.Label(Text="Ready - " + str(len(self.objects)) + " objects")
    
    def create_layout(self):
        layout = forms.DynamicLayout()
        layout.Spacing = drawing.Size(5, 5)
        layout.AddRow(self.key_label)
        layout.AddRow(self.key_dropdown)
        layout.AddRow(None)
        layout.AddRow(self.default_color_label)
        layout.AddRow(self.default_color_picker)
        layout.AddRow(None)
        layout.AddRow(self.legend_label)
        layout.AddRow(self.legend_panel)
        layout.AddRow(None)
        layout.AddRow(self.warnings_label)
        layout.AddRow(self.warnings_text)
        layout.AddRow(None)
        
        button_layout = forms.DynamicLayout()
        button_layout.Spacing = drawing.Size(5, 5)
        button_layout.AddRow(self.update_button, self.select_button, self.close_button)
        layout.AddRow(button_layout)
        layout.AddRow(self.status_label)
        
        self.Content = layout
    
    def update_legend(self):
        if not self.active_key:
            return
        legend_layout = forms.DynamicLayout()
        legend_layout.Spacing = drawing.Size(5, 5)
        legend_layout.Padding = drawing.Padding(10)
        
        title = forms.Label(Text=self.active_key)
        title.Font = drawing.Font(drawing.SystemFont.Bold, 12)
        legend_layout.AddRow(title)
        legend_layout.AddRow(None)
        
        key_colors = self.color_map.get(self.active_key, {})
        for value in sorted(key_colors.keys()):
            color_rgb = key_colors[value]
            color_panel = forms.Panel()
            color_panel.Size = drawing.Size(30, 20)
            color_panel.BackgroundColor = drawing.Color.FromArgb(color_rgb[0], color_rgb[1], color_rgb[2])
            value_label = forms.Label(Text=value)
            rgb_text = "RGB: {},{},{}".format(color_rgb[0], color_rgb[1], color_rgb[2])
            rgb_label = forms.Label(Text=rgb_text)
            rgb_label.Font = drawing.Font(drawing.SystemFont.Default, 9)
            rgb_label.TextColor = drawing.Colors.Gray
            
            row_layout = forms.DynamicLayout()
            row_layout.Spacing = drawing.Size(5, 5)
            row_layout.AddRow(color_panel, value_label, rgb_label)
            legend_layout.AddRow(row_layout)
        
        self.legend_panel.Content = legend_layout
    
    def update_warnings(self, stats):
        warnings = []
        self.problem_objects = []
        
        if stats['missing_color_definition']:
            warnings.append("ERROR: Values in objects but missing from CSV:")
            warnings.append("=" * 50)
            for value, obj_list in stats['missing_color_definition']:
                warnings.append("  '{}' ({} objects)".format(value, len(obj_list)))
                self.problem_objects.extend(obj_list)
            warnings.append("")
        
        if stats['unused_color_definitions']:
            warnings.append("WARNING: Values in CSV but not in selection:")
            warnings.append("=" * 50)
            for value in stats['unused_color_definitions']:
                warnings.append("  '{}'".format(value))
            warnings.append("")
        
        if stats['default'] > 0:
            warnings.append("INFO: {} objects without '{}' key".format(stats['default'], self.active_key))
            warnings.append("")
        
        if not warnings:
            warnings.append("SUCCESS: All objects colored correctly!")
            warnings.append("  {} objects colored".format(stats['colored']))
        
        self.warnings_text.Text = "\n".join(warnings)
        self.select_button.Enabled = len(self.problem_objects) > 0
    
    def on_key_changed(self, sender, e):
        selected = self.key_dropdown.SelectedValue
        if selected:
            self.active_key = str(selected)
            self.update_legend()
            self.warnings_text.Text = "Click 'Update Colors' to analyze..."
            self.select_button.Enabled = False
    
    def on_default_color_changed(self, sender, e):
        color = self.default_color_picker.Value
        self.default_color = (color.R, color.G, color.B)
    
    def on_update_clicked(self, sender, e):
        if not self.active_key:
            return
        
        self.status_label.Text = "Applying colors..."
        self.load_color_map()
        self.update_legend()
        
        stats = apply_colors_to_objects(self.objects, self.active_key, self.color_map, self.default_color)
        sc.doc.Views.Redraw()
        
        self.update_warnings(stats)
        self.status_label.Text = "Updated! Colored: {}, Default: {}".format(stats['colored'], stats['default'])
    
    def on_select_clicked(self, sender, e):
        if self.problem_objects:
            rs.UnselectAllObjects()
            rs.SelectObjects(self.problem_objects)
            forms.MessageBox.Show("Selected {} objects with missing colors".format(len(self.problem_objects)), "Objects Selected")
    
    def on_close_clicked(self, sender, e):
        self.Close(True)


# ============================================================================
# MAIN UNIFIED INTERFACE
# ============================================================================

class DataVisualizationTool(forms.Dialog[bool]):
    """Main unified interface - stays on top"""
    
    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.csv_folder = os.path.join(self.script_dir, "_CSVOutput")
        
        self.Title = "Data Visualization Tool"
        self.Padding = drawing.Padding(15)
        self.Resizable = False
        self.Size = drawing.Size(450, 420)
        self.Topmost = True  # KEEP ON TOP FIX
        
        self.create_controls()
        self.create_layout()
        
    def create_controls(self):
        self.header = forms.Label(Text="Rhino Metadata Visualization")
        self.header.Font = drawing.Font(drawing.SystemFont.Bold, 12)
        
        self.step1_label = forms.Label(Text="Step 1: Initialize Keys")
        self.step1_label.Font = drawing.Font(drawing.SystemFont.Bold, 10)
        self.step1_desc = forms.Label(Text="Apply metadata keys from CSV to objects")
        self.step1_desc.TextColor = drawing.Colors.Gray
        self.step1_button = forms.Button(Text="Select CSV & Apply Keys")
        self.step1_button.Click += self.on_step1
        
        self.step2_label = forms.Label(Text="Step 2: Extract Unique Values")
        self.step2_label.Font = drawing.Font(drawing.SystemFont.Bold, 10)
        self.step2_desc = forms.Label(Text="Scan objects and generate color CSV")
        self.step2_desc.TextColor = drawing.Colors.Gray
        self.step2_button = forms.Button(Text="Scan & Generate CSV")
        self.step2_button.Click += self.on_step2
        
        self.step3_label = forms.Label(Text="Step 3: Visualize with Colors")
        self.step3_label.Font = drawing.Font(drawing.SystemFont.Bold, 10)
        self.step3_desc = forms.Label(Text="Interactive color visualization")
        self.step3_desc.TextColor = drawing.Colors.Gray
        self.step3_button = forms.Button(Text="Open Color Manager")
        self.step3_button.Click += self.on_step3
        
        self.status = forms.Label(Text="Ready")
        self.status.TextColor = drawing.Colors.Blue
        self.close_button = forms.Button(Text="Close")
        self.close_button.Click += self.on_close
        
    def create_layout(self):
        layout = forms.DynamicLayout()
        layout.Spacing = drawing.Size(8, 8)
        layout.AddRow(self.header)
        layout.AddRow(None)
        layout.AddRow(self.step1_label)
        layout.AddRow(self.step1_desc)
        layout.AddRow(self.step1_button)
        layout.AddRow(None)
        layout.AddRow(self.step2_label)
        layout.AddRow(self.step2_desc)
        layout.AddRow(self.step2_button)
        layout.AddRow(None)
        layout.AddRow(self.step3_label)
        layout.AddRow(self.step3_desc)
        layout.AddRow(self.step3_button)
        layout.AddRow(None)
        layout.AddRow(self.status)
        layout.AddRow(self.close_button)
        self.Content = layout
        
    def on_step1(self, sender, e):
        csv_path = rs.OpenFileName("Select CSV", "CSV Files (*.csv)|*.csv||", folder=self.csv_folder if os.path.exists(self.csv_folder) else None)
        if not csv_path:
            return
        objects = rs.GetObjects("Select objects", preselect=True)
        if not objects:
            return
        keys = read_keys_from_csv(csv_path)
        if not keys:
            forms.MessageBox.Show("No keys in CSV", "Error")
            return
        stats = apply_keys_to_objects(objects, keys)
        self.status.Text = "Step 1: {} keys added".format(stats['keys_added'])
        forms.MessageBox.Show("Keys applied!\n\nObjects: {}\nKeys added: {}\nSkipped: {}".format(stats['objects_processed'], stats['keys_added'], stats['keys_skipped']), "Complete")
        self.BringToFront()  # Return focus
        
    def on_step2(self, sender, e):
        objects = rs.GetObjects("Select objects", preselect=True)
        if not objects:
            return
        kvd = collect_unique_values(objects)
        if not kvd:
            forms.MessageBox.Show("No metadata found", "Error")
            return
        csv_path = rs.SaveFileName("Save CSV", "CSV Files (*.csv)|*.csv||", folder=self.csv_folder if os.path.exists(self.csv_folder) else None, filename="UniqueValuesColorSettings.csv")
        if not csv_path:
            return
        rows = write_color_map_csv(csv_path, kvd)
        self.status.Text = "Step 2: {} values exported".format(rows)
        forms.MessageBox.Show("CSV generated!\n\nKeys: {}\nValues: {}\n\nNext: Fill RGB colors".format(len(kvd), rows), "Complete")
        self.BringToFront()  # Return focus
        
    def on_step3(self, sender, e):
        objects = rs.GetObjects("Select objects", preselect=True)
        if not objects:
            return
        csv_path = rs.OpenFileName("Select color CSV", "CSV Files (*.csv)|*.csv||", folder=self.csv_folder if os.path.exists(self.csv_folder) else None)
        if not csv_path:
            return
        # Launch Color Manager (child dialog)
        color_dlg = ColorManagerDialog(csv_path, objects)
        color_dlg.ShowModal(self)  # Modal to parent
        self.status.Text = "Step 3: Complete"
        self.BringToFront()  # Return focus
        
    def on_close(self, sender, e):
        self.Close(True)


def main():
    dialog = DataVisualizationTool()
    dialog.ShowModal(None)


if __name__ == "__main__":
    main()
