
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
        
        
integra_file_maker(df, os.getcwd())