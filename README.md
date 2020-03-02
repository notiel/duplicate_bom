# duplicate_bom
script for finding similar positions in xlsx BOMs

Use find_duplicates.py as main script: 
# for comparing two BOMS by PN
find_duplicates.py --compare BOM1.xlsx BOM2.xlsx 
# for detailed BOM comparing
find_duplicates.py --detailed BOM1.xlsx BOM2.xlsx 
# for finding similar components in path:
find_duplicates.py --duplicates PATH_TO_BOMS 
