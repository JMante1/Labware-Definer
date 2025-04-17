import pandas as pd
import os
import pathlib

def mantis_file_maker(plate_dimensions_df, output_dir):
	for index, row in plate_dimensions_df.iterrows():
		column_space = ((row['plate_length_mm'] - row['offset_left_to_a1_mm'] - row['offset_right_to_final_well_mm']) - (row['num_columns'] * row['well_length_mm']) )/ (row['num_columns'] - 1)
		column_pitch = column_space + row['well_length_mm']
		# if column_pitch< 0:
		#     raise ValueError(f"You are an idiot because your column_pitch is negative: {column_pitch} for plate {row['plate_name']}")
		
		row_space = ((row['plate_width_mm'] - row['offset_top_to_a1_mm'] - row['offset_bottom_to_final_well_mm']) - (row['num_rows'] * row['well_width_mm'])) / (row['num_rows'] - 1) 
		row_pitch = row_space + row['well_width_mm']
		if row_pitch > row['plate_width_mm']:
			raise ValueError(f'Row pitch value larger than your plate width for plate {row["plate_name"]}')
		
		if row['well_shape'] == 'square': 
			shape = 'Rectangle'
		elif row['well_shape'] == 'circular':
			shape = 'Circle'
		else:
			raise ValueError(f"Well not 'square' or 'circular' for plate {row['plate_name']}")

		if 1 <= row['plate_rigidity_scale'] <= 3:
			clamp_level = '1'
		elif 3< row['plate_rigidity_scale'] <= 6:
			clamp_level = '2'
		elif 6< row['plate_rigidity_scale'] <= 10:
			clamp_level =  '3'
		else:
			raise ValueError(f'Plate rigidity value is out of range for {row["plate_name"]}')
		
		z_mm = -(row['plate_height_mm'] + (1.5-9.665))
		if z_mm > row['plate_height_mm']:
			pass

		plate_string1 = f"""[ Version: 3 ]"""
		plate_string2 = f"""{row['plate_name'].replace(" ", "%32")} {row['num_rows']} {row['num_columns']} {row_pitch} {column_pitch} {row['plate_height_mm']} SBS%32Adapter.ad.txt zdrop:0 velocity:100 minsd:0 clamplevel:{clamp_level}""" 
		plate_string3 = f"""Well 1 X {row['offset_left_to_a1_mm']+(row['well_length_mm']/2)} Y {row['offset_top_to_a1_mm']+(row['well_width_mm']/2)} Z {z_mm} {row['well_length_mm']} {row['well_width_mm']} {shape} Well"""

		output_name = os.path.join(output_dir, f"{row['plate_name']}.pd.txt")
		with open(output_name, "w") as text_file:
			plate_string = '\r\n'.join([plate_string1, plate_string2, plate_string3, "", ""])
			text_file.write(plate_string)

# cwd = pathlib.Path(__file__).parent.resolve()
# top_folder = os.path.split(cwd)[0]
# top_folder = os.path.split(top_folder)[0]
# df = pd.read_excel(os.path.join(top_folder, 'Example Data', 'Labware Excel Input.xlsx'), skiprows=1)
# mantis_file_maker(df, top_folder)