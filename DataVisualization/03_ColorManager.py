"""
Script 03: Color Manager (CSV Version)
Interactive GUI for visualizing Rhino objects by metadata values.

Features:
- Dropdown to select active metadata key
- Display color legend with all values
- Default color picker for undefined objects
- Update button to refresh from CSV
- Screenshot-friendly legend display
- Comprehensive error detection and reporting

Author: RhinoGuire
Version: 1.1
"""

import rhinoscriptsyntax as rs
import scriptcontext as sc
import System
import csv
import os

# Import Eto forms
import Eto.Drawing as drawing
import Eto.Forms as forms


def read_color_map_from_csv(csv_path):
    """
    Reads color mapping from CSV file.
    
    Args:
        csv_path: Path to UniqueValuesColorSettings.csv
        
    Returns:
        Dictionary: {key_name: {value: (r, g, b, a)}}
    """
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
                    # Parse RGB values
                    try:
                        r_val = int(r) if r else 0
                        g_val = int(g) if g else 0
                        b_val = int(b) if b else 0
                        a_val = int(a) if a else 255
                        
                        # Validate ranges
                        r_val = max(0, min(255, r_val))
                        g_val = max(0, min(255, g_val))
                        b_val = max(0, min(255, b_val))
                        a_val = max(0, min(255, a_val))
                        
                        if key not in color_map:
                            color_map[key] = {}
                        
                        color_map[key][value] = (r_val, g_val, b_val, a_val)
                        
                    except ValueError:
                        # Skip rows with invalid color values
                        continue
        
        return color_map
        
    except Exception as e:
        print("Error reading CSV file: " + str(e))
        return None


def analyze_objects_for_key(objects, active_key):
    """
    Analyzes objects to find all unique values for the active key.
    
    Args:
        objects: List of Rhino object GUIDs
        active_key: Name of the active metadata key
        
    Returns:
        Dictionary with value counts and object lists
    """
    value_objects = {}  # {value: [list of object GUIDs]}
    objects_without_key = []
    
    for obj in objects:
        value = rs.GetUserText(obj, active_key)
        
        if value is None or value == '':
            objects_without_key.append(obj)
        else:
            if value not in value_objects:
                value_objects[value] = []
            value_objects[value].append(obj)
    
    return {
        'value_objects': value_objects,
        'objects_without_key': objects_without_key
    }


def apply_colors_to_objects(objects, active_key, color_map, default_color):
    """
    Applies colors to objects based on their metadata values.
    
    Args:
        objects: List of Rhino object GUIDs
        active_key: Name of the active metadata key
        color_map: Dictionary of color mappings
        default_color: RGB tuple for objects without the key
        
    Returns:
        Statistics dictionary with detailed error tracking
    """
    stats = {
        'colored': 0,
        'default': 0,
        'missing_color_definition': [],  # [(value, [object_guids])]
        'unused_color_definitions': []   # [values in CSV but not in selection]
    }
    
    key_color_map = color_map.get(active_key, {})
    
    # Analyze objects first
    analysis = analyze_objects_for_key(objects, active_key)
    value_objects = analysis['value_objects']
    objects_without_key = analysis['objects_without_key']
    
    # Track which CSV values are actually used
    used_values = set()
    
    # Apply colors to objects with the key
    for value, obj_list in value_objects.items():
        used_values.add(value)
        color = key_color_map.get(value)
        
        if color:
            # Apply color to all objects with this value
            for obj in obj_list:
                rs.ObjectColor(obj, (color[0], color[1], color[2]))
                stats['colored'] += 1
        else:
            # Value exists but no color defined - use default and track
            for obj in obj_list:
                rs.ObjectColor(obj, default_color)
            stats['missing_color_definition'].append((value, obj_list))
    
    # Apply default color to objects without the key
    for obj in objects_without_key:
        rs.ObjectColor(obj, default_color)
        stats['default'] += 1
    
    # Find unused color definitions (values in CSV but not in selection)
    csv_values = set(key_color_map.keys())
    unused = csv_values - used_values
    stats['unused_color_definitions'] = sorted(list(unused))
    
    return stats


class ColorManagerDialog(forms.Dialog[bool]):
    """Eto Forms dialog for color management"""
    
    def __init__(self, csv_path, objects):
        self.csv_path = csv_path
        self.objects = objects
        self.color_map = None
        self.active_key = None
        self.default_color = (128, 128, 128)  # Default gray
        
        # Load color map
        self.load_color_map()
        
        # Initialize dialog
        self.Title = "Color Manager - Data Visualization"
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        self.Size = drawing.Size(500, 600)
        
        # Create controls
        self.create_controls()
        self.create_layout()
        
    def load_color_map(self):
        """Load color map from CSV"""
        self.color_map = read_color_map_from_csv(self.csv_path)
        
        if not self.color_map:
            forms.MessageBox.Show("Error loading color map from CSV file.", "Error")
            self.Close(False)
    
    def create_controls(self):
        """Create all UI controls"""
        
        # Key selection dropdown
        self.key_label = forms.Label(Text="Select Metadata Key:")
        self.key_dropdown = forms.DropDown()
        
        # Populate dropdown with available keys
        available_keys = sorted(self.color_map.keys())
        for key in available_keys:
            self.key_dropdown.Items.Add(key)
        
        if available_keys:
            self.key_dropdown.SelectedIndex = 0
            self.active_key = available_keys[0]
        
        self.key_dropdown.SelectedIndexChanged += self.on_key_changed
        
        # Default color picker
        self.default_color_label = forms.Label(Text="Default Color (for objects without key):")
        self.default_color_picker = forms.ColorPicker()
        self.default_color_picker.Value = drawing.Color.FromArgb(128, 128, 128)
        self.default_color_picker.ValueChanged += self.on_default_color_changed
        
        # Legend area
        self.legend_label = forms.Label(Text="Color Legend:")
        self.legend_panel = forms.Panel()
        self.legend_panel.Size = drawing.Size(450, 250)
        self.legend_panel.BackgroundColor = drawing.Colors.White
        
        # Create legend content
        self.update_legend()
        
        # Warnings/Errors area
        self.warnings_label = forms.Label(Text="Warnings & Errors:")
        self.warnings_text = forms.TextArea()
        self.warnings_text.ReadOnly = True
        self.warnings_text.Size = drawing.Size(450, 100)
        self.warnings_text.Text = "Click 'Update Colors' to see analysis..."
        
        # Update button
        self.update_button = forms.Button(Text="Update Colors")
        self.update_button.Click += self.on_update_clicked
        
        # Select objects button
        self.select_button = forms.Button(Text="Select Problem Objects")
        self.select_button.Click += self.on_select_clicked
        self.select_button.Enabled = False
        self.problem_objects = []  # Store objects with issues
        
        # Close button
        self.close_button = forms.Button(Text="Close")
        self.close_button.Click += self.on_close_clicked
        
        # Status label
        self.status_label = forms.Label(Text="Ready - " + str(len(self.objects)) + " objects selected")
    
    def create_layout(self):
        """Create dialog layout"""
        
        # Main layout
        layout = forms.DynamicLayout()
        layout.Spacing = drawing.Size(5, 5)
        
        # Key selection
        layout.AddRow(self.key_label)
        layout.AddRow(self.key_dropdown)
        layout.AddRow(None)  # Spacer
        
        # Default color
        layout.AddRow(self.default_color_label)
        layout.AddRow(self.default_color_picker)
        layout.AddRow(None)  # Spacer
        
        # Legend
        layout.AddRow(self.legend_label)
        layout.AddRow(self.legend_panel)
        layout.AddRow(None)  # Spacer
        
        # Warnings
        layout.AddRow(self.warnings_label)
        layout.AddRow(self.warnings_text)
        layout.AddRow(None)  # Spacer
        
        # Buttons
        button_layout = forms.DynamicLayout()
        button_layout.Spacing = drawing.Size(5, 5)
        button_layout.AddRow(self.update_button, self.select_button, self.close_button)
        layout.AddRow(button_layout)
        
        # Status
        layout.AddRow(self.status_label)
        
        self.Content = layout
    
    def update_legend(self):
        """Update legend display with current key's values"""
        if not self.active_key:
            return
        
        # Create legend layout
        legend_layout = forms.DynamicLayout()
        legend_layout.Spacing = drawing.Size(5, 5)
        legend_layout.Padding = drawing.Padding(10)
        
        # Title
        title = forms.Label(Text=self.active_key)
        title.Font = drawing.Font(drawing.SystemFont.Bold, 12)
        legend_layout.AddRow(title)
        legend_layout.AddRow(None)  # Spacer
        
        # Get color map for this key
        key_colors = self.color_map.get(self.active_key, {})
        
        if not key_colors:
            no_colors_label = forms.Label(Text="No color definitions found for this key in CSV")
            no_colors_label.TextColor = drawing.Colors.Red
            legend_layout.AddRow(no_colors_label)
        else:
            # Add each value with its color
            for value in sorted(key_colors.keys()):
                color_rgb = key_colors[value]
                
                # Create color swatch
                color_panel = forms.Panel()
                color_panel.Size = drawing.Size(30, 20)
                color_panel.BackgroundColor = drawing.Color.FromArgb(
                    color_rgb[0], color_rgb[1], color_rgb[2]
                )
                
                # Create label
                value_label = forms.Label(Text=value)
                
                # Create RGB text
                rgb_text = "RGB: {},{},{}".format(color_rgb[0], color_rgb[1], color_rgb[2])
                rgb_label = forms.Label(Text=rgb_text)
                rgb_label.Font = drawing.Font(drawing.SystemFont.Default, 9)
                rgb_label.TextColor = drawing.Colors.Gray
                
                # Add row
                row_layout = forms.DynamicLayout()
                row_layout.Spacing = drawing.Size(5, 5)
                row_layout.AddRow(color_panel, value_label, rgb_label)
                
                legend_layout.AddRow(row_layout)
        
        # Set legend content
        self.legend_panel.Content = legend_layout
    
    def update_warnings(self, stats):
        """Update warnings text area with analysis results"""
        warnings = []
        self.problem_objects = []
        
        # Check for values without color definitions
        if stats['missing_color_definition']:
            warnings.append("ERROR: Values found in objects but missing from CSV:")
            warnings.append("=" * 50)
            for value, obj_list in stats['missing_color_definition']:
                warnings.append("  Value: '{}' ({} objects)".format(value, len(obj_list)))
                self.problem_objects.extend(obj_list)
            warnings.append("")
            warnings.append("These objects will use the default color.")
            warnings.append("Add these values to your CSV and fill in colors.")
            warnings.append("")
        
        # Check for unused color definitions
        if stats['unused_color_definitions']:
            warnings.append("WARNING: Values defined in CSV but not found in selection:")
            warnings.append("=" * 50)
            for value in stats['unused_color_definitions']:
                warnings.append("  Value: '{}'".format(value))
            warnings.append("")
            warnings.append("These color definitions are not being used.")
            warnings.append("Either the values don't exist in selected objects,")
            warnings.append("or you need to select more objects.")
            warnings.append("")
        
        # Check for objects without the key
        if stats['default'] > 0:
            warnings.append("INFO: Objects without '{}' key:".format(self.active_key))
            warnings.append("=" * 50)
            warnings.append("  {} objects don't have this metadata key".format(stats['default']))
            warnings.append("  These objects will use the default color.")
            warnings.append("")
        
        # Success message if no issues
        if not warnings:
            warnings.append("SUCCESS: All objects colored correctly!")
            warnings.append("=" * 50)
            warnings.append("  {} objects colored from CSV definitions".format(stats['colored']))
            warnings.append("  No errors or warnings detected.")
        
        self.warnings_text.Text = "\n".join(warnings)
        
        # Enable/disable select button based on whether there are problem objects
        self.select_button.Enabled = len(self.problem_objects) > 0
    
    def on_key_changed(self, sender, e):
        """Handle key selection change"""
        selected_item = self.key_dropdown.SelectedValue
        if selected_item:
            self.active_key = str(selected_item)
            self.update_legend()
            # Clear warnings when key changes
            self.warnings_text.Text = "Click 'Update Colors' to see analysis..."
            self.select_button.Enabled = False
    
    def on_default_color_changed(self, sender, e):
        """Handle default color change"""
        color = self.default_color_picker.Value
        self.default_color = (color.R, color.G, color.B)
    
    def on_update_clicked(self, sender, e):
        """Handle update button click"""
        if not self.active_key:
            forms.MessageBox.Show("Please select a metadata key.", "Error")
            return
        
        # Reload color map from CSV
        self.status_label.Text = "Reloading color map..."
        self.load_color_map()
        self.update_legend()
        
        # Apply colors
        self.status_label.Text = "Applying colors..."
        stats = apply_colors_to_objects(
            self.objects,
            self.active_key,
            self.color_map,
            self.default_color
        )
        
        # Redraw
        sc.doc.Views.Redraw()
        
        # Update warnings
        self.update_warnings(stats)
        
        # Update status
        status_text = "Updated! Colored: {}, Default: {}, Missing: {}".format(
            stats['colored'],
            stats['default'],
            len(stats['missing_color_definition'])
        )
        self.status_label.Text = status_text
        
        # Console output
        print("\n" + "="*60)
        print("COLOR UPDATE COMPLETE")
        print("="*60)
        print("Active key: " + self.active_key)
        print("Objects colored successfully: " + str(stats['colored']))
        print("Objects using default color: " + str(stats['default']))
        
        if stats['missing_color_definition']:
            print("\nERROR: Values without color definitions:")
            for value, obj_list in stats['missing_color_definition']:
                print("  '{}': {} objects".format(value, len(obj_list)))
        
        if stats['unused_color_definitions']:
            print("\nWARNING: Unused color definitions:")
            for value in stats['unused_color_definitions']:
                print("  '{}'".format(value))
    
    def on_select_clicked(self, sender, e):
        """Handle select problem objects button click"""
        if self.problem_objects:
            rs.UnselectAllObjects()
            rs.SelectObjects(self.problem_objects)
            message = "Selected {} objects with missing color definitions.\n\nThese objects have values not defined in your CSV.".format(
                len(self.problem_objects)
            )
            forms.MessageBox.Show(message, "Objects Selected")
    
    def on_close_clicked(self, sender, e):
        """Handle close button click"""
        self.Close(True)


def main():
    """Main execution function"""
    
    # Select objects
    objects = rs.GetObjects(
        "Select objects to visualize (colors will be applied to these only)",
        preselect=True
    )
    
    if not objects:
        print("No objects selected. Operation cancelled.")
        return
    
    print("Selected " + str(len(objects)) + " objects for visualization")
    
    # Auto-detect _CSVOutput folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_folder = os.path.join(script_dir, "_CSVOutput")
    
    # Get CSV file path
    csv_path = rs.OpenFileName(
        "Select UniqueValuesColorSettings.csv file",
        "CSV Files (*.csv)|*.csv||",
        folder=csv_folder if os.path.exists(csv_folder) else None
    )
    
    if not csv_path:
        print("No file selected. Operation cancelled.")
        return
    
    # Show dialog
    dialog = ColorManagerDialog(csv_path, objects)
    dialog.ShowModal(None)


# Run the script
if __name__ == "__main__":
    main()
