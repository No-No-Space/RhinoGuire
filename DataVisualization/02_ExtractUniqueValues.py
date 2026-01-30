"""
Script 02: Extract Unique Values (CSV Version)
Scans selected Rhino objects for metadata keys and generates color mapping CSV.

OUTPUT: Creates/updates UniqueValuesColorSettings.csv with format:
Key,Value,R,G,B,A
Building Year,1950,,,
Building Year,1975,,,
Condition,Good,,,

Author: RhinoGuire
Version: 1.1 (CSV)
"""

import rhinoscriptsyntax as rs
import csv
from collections import defaultdict
import os


def collect_unique_values(objects):
    """
    Scans objects and collects all unique key-value pairs.
    
    Args:
        objects: List of Rhino object GUIDs
        
    Returns:
        Dictionary: {key_name: set(unique_values)}
    """
    key_values = defaultdict(set)
    
    for obj in objects:
        # Get all user text keys for this object
        keys = rs.GetUserText(obj)
        
        if keys:
            for key in keys:
                value = rs.GetUserText(obj, key)
                if value is not None and value != '':
                    key_values[key].add(value)
    
    # Convert sets to sorted lists
    return {key: sorted(list(values)) for key, values in key_values.items()}


def read_existing_colors(csv_path):
    """
    Reads existing color assignments from CSV to preserve user's work.
    
    Args:
        csv_path: Path to existing CSV file
        
    Returns:
        Dictionary: {(key, value): (r, g, b, a)}
    """
    existing_colors = {}
    
    if not os.path.exists(csv_path):
        return existing_colors
    
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
                    existing_colors[(key, value)] = (r, g, b, a)
    except Exception as e:
        print("Note: Could not read existing colors: " + str(e))
    
    return existing_colors


def write_color_map_csv(csv_path, key_value_dict):
    """
    Writes color mapping data to CSV file.
    Preserves existing color assignments if they exist.
    
    Args:
        csv_path: Path to output CSV file
        key_value_dict: Dictionary of {key_name: [unique_values]}
        
    Returns:
        Number of rows written
    """
    # Read existing color assignments
    existing_colors = read_existing_colors(csv_path)
    
    rows_written = 0
    
    try:
        with open(csv_path, 'w') as file:
            writer = csv.writer(file)
            
            # Write header
            writer.writerow(['Key', 'Value', 'R', 'G', 'B', 'A'])
            
            # Write data rows, sorted by key name
            for key_name in sorted(key_value_dict.keys()):
                unique_values = key_value_dict[key_name]
                
                for value in unique_values:
                    # Check if color already exists for this key-value pair
                    existing_color = existing_colors.get((key_name, value), ('', '', '', ''))
                    
                    writer.writerow([
                        key_name,
                        value,
                        existing_color[0],
                        existing_color[1],
                        existing_color[2],
                        existing_color[3]
                    ])
                    rows_written += 1
        
        return rows_written
        
    except Exception as e:
        print("Error writing CSV file: " + str(e))
        return 0


def main():
    """Main execution function"""
    
    # Select objects to scan
    objects = rs.GetObjects(
        "Select objects to scan for metadata keys",
        preselect=True
    )
    
    if not objects:
        print("No objects selected. Operation cancelled.")
        return
    
    print("Scanning " + str(len(objects)) + " objects for metadata...")
    
    # Collect unique values
    key_value_dict = collect_unique_values(objects)
    
    if not key_value_dict:
        print("\nNo metadata keys found on selected objects.")
        rs.MessageBox(
            "No metadata keys found on the selected objects.\n\n"
            "Make sure you've run Script 01 first to add keys,\n"
            "and filled in some values using SetUserText.",
            0, "No Metadata Found"
        )
        return
    
    print("\nFound " + str(len(key_value_dict)) + " unique keys:")
    total_values = 0
    for key, values in sorted(key_value_dict.items()):
        print("  - " + key + ": " + str(len(values)) + " unique values")
        total_values += len(values)
    
    # Auto-detect _CSVOutput folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_folder = os.path.join(script_dir, "_CSVOutput")
    
    # Get CSV file path for output
    csv_path = rs.SaveFileName(
        "Save color mapping CSV as",
        "CSV Files (*.csv)|*.csv||",
        folder=csv_folder if os.path.exists(csv_folder) else None,
        filename="UniqueValuesColorSettings.csv"
    )
    
    if not csv_path:
        print("\nNo file selected. Operation cancelled.")
        return
    
    # Write CSV file
    print("\nWriting CSV file: " + csv_path)
    rows_written = write_color_map_csv(csv_path, key_value_dict)
    
    # Report results
    print("\n" + "="*60)
    print("OPERATION COMPLETE")
    print("="*60)
    print("Objects scanned: " + str(len(objects)))
    print("Unique keys found: " + str(len(key_value_dict)))
    print("Total unique values: " + str(total_values))
    print("Rows written to CSV: " + str(rows_written))
    print("\nNext steps:")
    print("1. Open the CSV file (or import to Excel)")
    print("2. Fill in R, G, B, A values for each unique value")
    print("3. Save the CSV")
    print("4. Run Script 03 to visualize objects with these colors")
    
    # Show summary dialog
    message = (
        "Color mapping CSV generated successfully!\n\n"
        "Objects scanned: " + str(len(objects)) + "\n"
        "Keys found: " + str(len(key_value_dict)) + "\n"
        "Unique values: " + str(total_values) + "\n\n"
        "CSV Format:\n"
        "Key, Value, R, G, B, A\n\n"
        "Next: Fill in RGB values in the CSV,\n"
        "then run Script 03 to apply colors."
    )
    rs.MessageBox(message, 0, "Extract Unique Values - Complete")


# Run the script
if __name__ == "__main__":
    main()
