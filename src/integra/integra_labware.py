
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
            bottom_shape = 'Square'
        elif row['well_bottom_shape'] == 'U-Bottom':
            bottom_shape = 'UShape'
        elif row['well_bottom_shape'] == 'V-Bottom':
            bottom_shape = 'VShape'
            v_shape_depth = (row['well_max_depth_mm'] - row['bottom_start_depth_mm']) * 100
        else:
            raise ValueError(f"Well not 'flat', 'u-bottom', or 'v-bottom' for plate {row['plate_name']}")
            
        
        first_hole_position_text = f"{round((row['offset_left_to_a1_mm']+ row['well_length_mm']) * 100)};{round((row['offset_top_to_a1_mm'] + row['well_width_mm'])* 100)}"

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
        


        root = ET.Element('Plate')
        root.attrib['xmlns:xsd'] = "http://www.w3.org/2001/XMLSchema"
        root.attrib['xmlns:xsi'] = "http://www.w3.org/2001/XMLSchema-instance"
        root.attrib['Version'] = "1"

        data_plate = ET.Element('DataVersion') 
        root.append (data_plate)
        data_plate.text = str(0) 

        name_plate = ET.Element('Name')
        root.append(name_plate)
        name_plate.text = row['plate_name']

        manu_plate = ET.Element('Manufacturer') 
        root.append (manu_plate)
        manu_plate.text = row['manufacturer']

        part_plate = ET.Element('PartNumber')
        root.append(part_plate)
        part_plate.text = str(row['catalog_number'])

        description_plate = ET.Element('Description')
        root.append(description_plate)
        description_plate.text = row['plate_name']

        measure_plate = ET.Element('Measurements') 
        measure_plate.attrib['Version'] = "0"
        root.append (measure_plate) 

        data_measure = ET.SubElement(measure_plate, 'DataVersion')
        data_measure.text = str(0) 

        description_measure = ET.SubElement(measure_plate, 'Description') 
      
        length_measure = ET.SubElement(measure_plate, 'FootprintLengthMM')
        length_measure.text = str(round(foot_print_length_mm))

        width_measure = ET.SubElement(measure_plate, 'FootprintWidthMM') 
        width_measure.text = str(round(foot_print_width_mm))

        height_measure = ET.SubElement(measure_plate, 'HeightMM')
        height_measure.text = str(round(height_mm))

        well_plate = ET.Element('Wells')
        well_plate.attrib['Version'] = "0"
        root.append(well_plate)

        data_well = ET.SubElement(well_plate, 'DataVersion') 
        data_well.text = str(0)

        description_well = ET.SubElement(well_plate, 'Description')
        
        angle_well = ET.SubElement(well_plate, 'Angle')
        angle_well.text = str(0)
         
        sectionheight_well = ET.SubElement(well_plate, 'SectionHeightCorrection')
        sectionheight_well.text = str(0)
        
        deltamax_well = ET.SubElement(well_plate, 'DeltaHmax') 
        deltamax_well.text = str(0)
         
        bottomshape_well = ET.SubElement(well_plate, 'BottomShape')
        bottomshape_well.text = bottom_shape

        columncount_well = ET.SubElement(well_plate, 'CollumnCount')
        columncount_well.text = str(row['num_columns'])
        
        columngap_well = ET.SubElement(well_plate, 'CollumnGap') 
        columngap_well.text = str(round(column_gap))
         
        depth_well = ET.SubElement(well_plate, 'Depth')
        depth_well.text = str(round(depth))
        
        volume_well = ET.SubElement(well_plate, 'NominalWellVolume')
        volume_well.text = str(round(row['max_volume_ul'] * 100))

        if row['well_bottom_shape'] == 'V-Bottom':
            shapedepth_well = ET.SubElement(well_plate, 'VShapeDepth')
            shapedepth_well.text = str(round(v_shape_depth))
        
        firsthole = ET.SubElement(well_plate, 'FirstHolePositionText') 
        firsthole.text = str(first_hole_position_text)
         
        rows = ET.SubElement(well_plate, 'RowCount')
        rows.text = str(row['num_rows'])
        
        rowgap = ET.SubElement(well_plate, 'RowGap')
        rowgap.text = str(round(row_gap)) 

        shape_well = ET.SubElement(well_plate, 'Shape')
        shape_well.text = shape

        size_well = ET.SubElement(well_plate, 'Size') 
        size_well.text = str(round(size))

        length_well = ET.SubElement(well_plate, 'Length')
        length_well.text = str(round(length))
        
        sizebottom_well = ET.SubElement(well_plate, 'SizeBottom') 
        sizebottom_well.text = str(0)
         

        tree = ET.ElementTree(root) 
        ET.indent(tree, space="\t", level=0)
        # tree.write(os.path.join(output_dir, f"output2.xml"), encoding="utf-8", xml_declaration=True)
        tree.write(os.path.join(output_dir, f"{row['catalog_number']}.xml"), encoding="utf-8", xml_declaration=True)

integra_file_maker(df, os.getcwd())