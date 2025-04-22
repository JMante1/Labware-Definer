import os
import xml.etree.ElementTree as ET


def create_and_append(parent, tag, text):
        el = ET.SubElement(parent, tag)
        el.text = str(text)
        return el

def integra_file_maker_optimized(plate_dimensions_df, output_dir):
    bottom_shape_map = {
        'Flat': 'Square',
        'U-Bottom': 'UShape',
        'V-Bottom': 'VShape'
    }

    well_shape_map = {
        'square': 'Square',
        'circular': 'Circle'
    }

    for _, row in plate_dimensions_df.iterrows():
        mm_to_um = lambda x: x * 100
        round_mm = lambda x: round(x * 100)

        # Preprocess values
        bottom_shape = bottom_shape_map.get(row['well_bottom_shape'])
        if not bottom_shape:
            raise ValueError(f"Invalid well bottom shape: {row['well_bottom_shape']}")

        shape = well_shape_map.get(row['well_shape'])
        if not shape:
            raise ValueError(f"Invalid well shape: {row['well_shape']}")

        if bottom_shape == 'VShape':
            v_shape_depth = round_mm(row['well_max_depth_mm'] - row['bottom_start_depth_mm'])

        footprint_length = round_mm(row['plate_length_mm'])
        footprint_width = round_mm(row['plate_width_mm'])
        height = round_mm(row['plate_height_mm'])
        depth = round_mm(row['well_max_depth_mm'])
        column_gap = round_mm(
            ((row['plate_length_mm'] - row['offset_left_to_a1_mm'] - row['offset_right_to_final_well_mm']) -
             row['num_columns'] * row['well_length_mm']) / (row['num_columns'] - 1) + row['well_length_mm'])
        row_gap = round_mm(
            ((row['plate_width_mm'] - row['offset_top_to_a1_mm'] - row['offset_bottom_to_final_well_mm']) -
             row['num_rows'] * row['well_width_mm']) / (row['num_rows'] - 1) + row['well_width_mm'])
        size = round_mm(row['well_length_mm'])
        length = round_mm(row['well_width_mm'])
        first_hole_position_text = f"{round_mm(row['offset_left_to_a1_mm'])};{round_mm(row['offset_top_to_a1_mm'])}"

        # XML generation
        root = ET.Element('Plate', {
            'xmlns:xsd': "http://www.w3.org/2001/XMLSchema",
            'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance",
            'Version': "1"
        })

        for tag, text in [
            ('DataVersion', 0),
            ('Name', row['plate_name']),
            ('Manufacturer', row['manufacturer']),
            ('PartNumber', row['catalog_number']),
            ('Description', row['plate_name'])
        ]:
            create_and_append(root, tag, text)

        measurements = ET.SubElement(root, 'Measurements', {'Version': "0"})
        for tag, text in [
            ('DataVersion', 0),
            ('Description', ''),
            ('FootprintLengthMM', footprint_length),
            ('FootprintWidthMM', footprint_width),
            ('HeightMM', height)
        ]:
            create_and_append(measurements, tag, text)

        wells = ET.SubElement(root, 'Wells', {'Version': "0"})
        for tag, text in [
            ('DataVersion', 0),
            ('Description', ''),
            ('Angle', 0),
            ('SectionHeightCorrection', 0),
            ('DeltaHmax', 0),
            ('BottomShape', bottom_shape),
            ('CollumnCount', row['num_columns']),
            ('CollumnGap', column_gap),
            ('Depth', depth),
            ('NominalWellVolume', round(row['max_volume_ul'] * 100)),
            ('FirstHolePositionText', first_hole_position_text),
            ('RowCount', row['num_rows']),
            ('RowGap', row_gap),
            ('Shape', shape),
            ('Size', size),
            ('Length', length),
            ('SizeBottom', 0)
        ]:
            create_and_append(wells, tag, text)

        if bottom_shape == 'VShape':
            create_and_append(wells, 'VShapeDepth', v_shape_depth)

        tree = ET.ElementTree(root)
        ET.indent(tree, space="\t", level=0)
        output_path = os.path.join(output_dir, f"{row['catalog_number']}.xml")
        tree.write(output_path, encoding="utf-8", xml_declaration=True)
