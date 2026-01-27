import rhinoscriptsyntax as rs

# Select objects
objs = rs.GetObjects("Select buildings")

# Check if selection was successful
if objs is None:
    print("No objects selected")
else:
    # Define your keys with default values
    keys = {
        "Building Year": "",
        "Latest Renovation": "",
        "Condition": ""
    }
    
    # Apply to all objects
    for obj in objs:
        for key, value in keys.items():
            rs.SetUserText(obj, key, value)
    
    print("Added keys to {} objects".format(len(objs)))