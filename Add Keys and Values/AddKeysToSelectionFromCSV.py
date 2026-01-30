import rhinoscriptsyntax as rs
import csv

# Get CSV file path
csv_path = rs.OpenFileName("Select CSV file with keys", "CSV Files (*.csv)|*.csv||")

if csv_path is None:
    print("No file selected")
else:
    # Read keys from CSV
    keys = {}
    with open(csv_path, 'r') as file:
        reader = csv.reader(file)
        next(reader, None)  # Skip header row if present
        for row in reader:
            if len(row) >= 2:
                key_name = row[0].strip()
                default_value = row[1].strip()
                keys[key_name] = default_value
    
    if not keys:
        print("No keys found in CSV")
    else:
        print("Found {} keys in CSV".format(len(keys)))
        
        # Select objects
        objs = rs.GetObjects("Select buildings to add keys to")
        
        if objs is None:
            print("No objects selected")
        else:
            # Apply keys to all objects
            for obj in objs:
                for key, value in keys.items():
                    # Check if key already exists
                    existing = rs.GetUserText(obj, key)
                    if existing is None:
                        # Key doesn't exist, add it
                        rs.SetUserText(obj, key, value)
                    else:
                        # Key exists, skip
                        print("Key '{}' already exists on object, skipping".format(key))
            
            print("Finished processing {} objects".format(len(objs)))