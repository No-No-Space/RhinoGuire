"""
Script 01: Initialize Keys (CSV Version)
Reads key definitions from CSV and applies them to selected Rhino objects.
Preserves existing keys and values to avoid data loss.

IMPORTANT: Before running this script, export your Excel table to CSV format
and place it in the _CSVOutput folder.

Author: RhinoGuire
Version: 1.1 (CSV)
"""

import rhinoscriptsyntax as rs
import csv
import os


def read_keys_from_csv(csv_path):
    """
    Reads key definitions from CSV file.
    
    Expected CSV format:
    Key Name,Default Value
    Building Year,
    Condition,Good
    
    Args:
        csv_path: Path to the CSV file
        
    Returns:
        Dictionary of {key_name: default_value}
    """
    try:
        keys = {}
        
        with open(csv_path, 'r') as file:
            reader = csv.reader(file)
            
            # Skip header row
            header = next(reader, None)
            if not header:
                print("ERROR: CSV file is empty")
                return None
            
            # Read key definitions
            for row in reader:
                if len(row) >= 2:
                    key_name = row[0].strip()
                    default_value = row[1].strip() if row[1] else ''
                    
                    if key_name:  # Skip empty rows
                        keys[key_name] = default_value
        
        return keys
        
    except Exception as e:
        print("Error reading CSV file: " + str(e))
        return None


def apply_keys_to_objects(objects, keys):
    """
    Applies keys to Rhino objects without overwriting existing keys/values.
    
    Args:
        objects: List of Rhino object GUIDs
        keys: Dictionary of {key_name: default_value}
        
    Returns:
        Statistics dictionary
    """
    stats = {
        'objects_processed': 0,
        'keys_added': 0,
        'keys_skipped': 0
    }
    
    for obj in objects:
        stats['objects_processed'] += 1
        
        for key_name, default_value in keys.items():
            # Check if key already exists
            existing_value = rs.GetUserText(obj, key_name)
            
            if existing_value is None:
                # Key doesn't exist, add it with default value
                rs.SetUserText(obj, key_name, default_value)
                stats['keys_added'] += 1
            else:
                # Key exists, skip to preserve existing data
                stats['keys_skipped'] += 1
    
    return stats


def main():
    """Main execution function"""
    
    # Auto-detect _CSVOutput folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_folder = os.path.join(script_dir, "_CSVOutput")
    
    # Get CSV file path (default to _CSVOutput folder)
    csv_path = rs.OpenFileName(
        "Select CSV file (exported from Excel)",
        "CSV Files (*.csv)|*.csv||",
        folder=csv_folder if os.path.exists(csv_folder) else None
    )
    
    if not csv_path:
        print("No file selected. Operation cancelled.")
        return
    
    # Read keys from CSV
    print("Reading keys from: " + csv_path)
    keys = read_keys_from_csv(csv_path)
    
    if not keys:
        print("No keys found in CSV file or error reading file.")
        rs.MessageBox(
            "No keys found in the CSV file.\n\n"
            "Make sure you exported the Excel table correctly:\n"
            "- First row: Key Name,Default Value\n"
            "- Following rows: Your key definitions",
            0, "Error"
        )
        return
    
    print("\nFound " + str(len(keys)) + " keys:")
    for key_name, default_value in keys.items():
        default_display = "'" + default_value + "'" if default_value else "(empty)"
        print("  - " + key_name + ": " + default_display)
    
    # Select objects
    objects = rs.GetObjects(
        "Select objects to apply " + str(len(keys)) + " metadata keys",
        preselect=True
    )
    
    if not objects:
        print("\nNo objects selected. Operation cancelled.")
        return
    
    # Apply keys
    print("\nApplying keys to " + str(len(objects)) + " objects...")
    stats = apply_keys_to_objects(objects, keys)
    
    # Report results
    print("\n" + "="*50)
    print("OPERATION COMPLETE")
    print("="*50)
    print("Objects processed: " + str(stats['objects_processed']))
    print("Keys added: " + str(stats['keys_added']))
    print("Keys skipped (already existed): " + str(stats['keys_skipped']))
    print("\nExisting keys and values were preserved.")
    
    # Show summary dialog
    message = (
        "Keys applied successfully!\n\n"
        "Objects processed: " + str(stats['objects_processed']) + "\n"
        "Keys added: " + str(stats['keys_added']) + "\n"
        "Keys skipped: " + str(stats['keys_skipped']) + "\n\n"
        "Existing keys and values were preserved."
    )
    rs.MessageBox(message, 0, "Initialize Keys - Complete")


# Run the script
if __name__ == "__main__":
    main()
