# duplicate_bom
script for finding similar positions in xlsx BOMs

Use find_duplicates.py as main script 
Use find_similar() as main function to find similar positions in one .xlsx BOM file or folder with BOM files. Use file or folder as script parameter. 
Use compare_boms() to compare two BOM files (given as script parameters). Only added/deleted components will be shown. 
USe compare_boms(detailed=True) to compare two BOM files and see all differences
BOM files should have PN, designator, value, component type, footprint fields
