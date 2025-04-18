
from logging import root
import pandas as pd
import pathlib
import xml.etree.ElementTree as ET
import os

cwd = pathlib.Path(__file__).parent.resolve()
top_folder = os.path.split(cwd)[0]
top_folder = os.path.split(top_folder)[0]
df = pd.read_excel(os.path.join(top_folder, 'Example Data', 'Labware Excel Input.xlsx'), skiprows=1)

def integra_file_maker(plate_dimensions_df, output_dir):
    for index, row in plate_dimensions_df.iterrows():
        foot_print_length_mm = row['plate_length_mm'] * 100
        
        foot_print_width_mm = row['plate_width_mm'] * 100

        height_mm = row['plate_height_mm'] * 100

        column_gap = ((((row['plate_length_mm'] - row['offset_left_to_a1_mm'] - row['offset_right_to_final_well_mm']) - (row['num_columns'] * row['well_length_mm']) )/ (row['num_columns'] - 1)) + row['well_length_mm']) * 100

        row_gap = ((((row['plate_width_mm'] - row['offset_top_to_a1_mm'] - row['offset_bottom_to_final_well_mm']) - (row['num_rows'] * row['well_width_mm'])) / (row['num_rows'] - 1)) + row['well_width_mm']) * 100

        depth = row['well_max_depth_mm'] * 100

        if row['well_bottom_shape'] == 'Flat': 
            shape = 'Square'
        elif row['well_bottom_shape'] == 'U-Bottom':
            shape = 'UShape'
        elif row['well_bottom_shape'] == 'V-Bottom':
            shape = 'VShape'
            v_shape_depth = (row['well_max_depth_mm'] - row['bottom_start_depth_mm']) * 100
        else:
            raise ValueError(f"Well not 'flat', 'u-bottom', or 'v-bottom' for plate {row['plate_name']}")
            
        
        first_hole_position_text = row['offset_left_to_a1_mm'] * 100

        if row['well_shape'] == 'square': 
            shape = 'Square'
        elif row['well_shape'] == 'circular':
            shape = 'Circle'
        else:
            raise ValueError(f"Well not 'square' or 'circular' for plate {row['plate_name']}")
     
        size = row['well_length_mm'] * 100
     
        length = row['well_width_mm'] * 100

        size_bottom = 0
        if size_bottom != 0:
          raise ValueError(f'always zero')
        
        # create output file
        
import xml.etree.ElementTree as ET  

def GenerateXML(fileName) : 

    root = ET.Element('Plate')
    root.attrib['xmlns:xsd'] = "http://www.w3.org/2001/XMLSchema"
    root.attrib['xmlns:xsi'] = "http://www.w3.org/2001/XMLSchema-instance"
    root.attrib['Version'] = "1"

    data_plate = ET.Element('DataVersion') 
    root.append (data_plate) 

    name_plate = ET.Element('Name')
    root.append(name_plate)

    manu_plate = ET.Element('Manufacturer') 
    root.append (manu_plate) 

    part_plate = ET.Element('PartNumber')
    root.append(part_plate)

    description_plate = ET.Element('Description')
    root.append(description_plate)

    measure_plate = ET.Element('Measurements') 
    measure_plate.attrib['Version'] = "0"
    root.append (measure_plate) 

    data_measure = ET.SubElement(measure_plate, 'DataVersion')
    root.append(data_measure)

    description_measure = ET.SubElement(measure_plate, 'Description') 
    root.append (data_measure) 

    length_measure = ET.SubElement(measure_plate, 'FootprintLengthMM')
    root.append(length_measure)

    width_measure = ET.SubElement(measure_plate, 'FootprintWidthMM') 
    root.append (width_measure) 

    height_measure = ET.SubElement(measure_plate, 'HeightMM')
    root.append(height_measure)

    m11 = ET.Element('Measurements') 
    root.append (m11) 

    m12 = ET.Element('Wells')
    m12.attrib['Version'] = "0"
    root.append(m12)

    m13 = ET.SubElement(m12, 'DataVersion') 
    root.append (m13) 

    m14 = ET.SubElement(m12, 'Description')
    root.append(m14)

    m15 = ET.SubElement(m12, 'Angle') 
    root.append (m15) 

    m16 = ET.SubElement(m12, 'SectionHeightCorrection')
    root.append(m16)

    m17 = ET.SubElement(m12, 'DeltaHmax') 
    root.append (m17) 

    m18 = ET.SubElement(m12, 'ColumnCount')
    root.append(m18)

    m19 = ET.SubElement(m12, 'CollumnGap') 
    root.append (m19) 

    m20 = ET.SubElement(m12, 'Depth')
    root.append(m20)

    m21 = ET.SubElement(m12, 'NominalWellVolume') 
    root.append (m21) 

    m22 = ET.SubElement(m12, 'VShapeDepth')
    root.append(m22)

    m23 = ET.SubElement(m12, 'FirstHolePositionText') 
    root.append (m23) 

    m25 = ET.SubElement(m12, 'RowCount')
    root.append(m25)

    m26 = ET.SubElement(m12, 'RowGap') 
    root.append (m26) 

    m27 = ET.SubElement(m12, 'Shape')
    root.append(m27)

    m28 = ET.SubElement(m12, 'Size') 
    root.append (m28) 

    m29 = ET.SubElement(m12, 'Length')
    root.append(m29)

    m30 = ET.SubElement(m12, 'SizeBottom') 
    root.append (m30) 

    

    tree = ET.ElementTree(root) 
    ET.indent(tree, space="\t", level=0)
    tree.write(fileName, encoding="utf-8")

integra_file_maker(df, os.getcwd())
GenerateXML('output2.xml')