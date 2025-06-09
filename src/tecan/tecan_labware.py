import pyautogui
import os
import win32gui
import pytesseract
import cv2
import matplotlib.pyplot as plt
import pyperclip
import numpy as np
import time
import pandas as pd
import re

# must have tesseract installed from: https://github.com/tesseract-ocr/tesseract
# make sure info pad is closed
# may need to make sure the 'don't ask again' for delete is checked
# make sure fluent control is open and not minimised and settings>controlbar>power

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
tecan_dir = os.path.dirname(os.path.abspath(__file__))
screenshots_dir = os.path.join(tecan_dir, "screenshots")
# results_dir = os.path.join(tecan_dir, "results")
pyautogui.PAUSE = 0.5  

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if wildcard == str(win32gui.GetWindowText(hwnd)):
            self._handle = hwnd
            # print(str(win32gui.GetWindowText(hwnd)))

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)
 

def clean_string(s):
    """Remove non-alphanumeric characters and lowercase."""
    return re.sub(r'[^a-zA-Z0-9_]', '', s).lower()

def find_phrase(phrase, data):
    print(data)
    # Clean  and split the phrase into words
    phrase_words = [clean_string(word) for word in phrase.split()]
    print(phrase_words)

    n_boxes = len(data['text'])
    i = 0

    while i < n_boxes:
        # Clean the OCR word
        word = clean_string(data['text'][i])

        # Check if current word matches the first word of the phrase
        if word == phrase_words[0]:
            match = True
            for j in range(1, len(phrase_words)):
                if i + j >= n_boxes or clean_string(data['text'][i + j]) != phrase_words[j]:
                    match = False
                    break

            if match:
                # Pull bounding box info for the whole phrase
                x_start = data['left'][i]
                y_start = data['top'][i]
                x_end = data['left'][i + len(phrase_words) - 1] + data['width'][i + len(phrase_words) - 1]
                y_end = data['top'][i + len(phrase_words) - 1] + data['height'][i + len(phrase_words) - 1]

                # Center of the bounding box
                center_x = (x_start + x_end) // 2
                center_y = (y_start + y_end) // 2

                x_w = x_end - x_start
                y_h = y_end - y_start
                print(x_w, x_end, x_start, phrase)
                return {'left': x_start, 'top': y_start, 'width': x_w, 'height': y_h}

        i += 1

    # If not found
    return {}

def take_screenshot(ialpha=7, ibeta=0, alpha=1.5, beta=0, show_screenshot=False):
    screenshot = pyautogui.screenshot()
    scrnshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    scrnshot = cv2.cvtColor(scrnshot, cv2.COLOR_BGR2GRAY)
    scrnshot = cv2.convertScaleAbs(scrnshot, alpha=alpha, beta=beta)

    # invert image
    inv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    inv  = cv2.cvtColor(inv, cv2.COLOR_BGR2GRAY)
    inv = cv2.bitwise_not(inv)
    screenshot_inv = cv2.convertScaleAbs(inv, alpha=ialpha, beta=ibeta)

    if show_screenshot:
        cv2.imshow('screenshot', scrnshot)
        cv2.waitKey(10000)
        cv2.imshow('inverted screenshot', screenshot_inv)
        cv2.waitKey(10000)
    return scrnshot, screenshot_inv


def click_word_on_screen(
    find_word, expected_match_num=1, click_type='left',
    sort_key='top', reversed_sort=False, click_local='center', show_screenshot=False, ialpha=7, ibeta=0, alpha=1.5, beta=0):
    # take screenshot
    scrnshot, screenshot_inv = take_screenshot(ialpha, ibeta, alpha, beta, show_screenshot)

    # convert to text data
    data = pytesseract.image_to_data(scrnshot, output_type=pytesseract.Output.DICT)
    data_i = pytesseract.image_to_data(screenshot_inv, output_type=pytesseract.Output.DICT)

    if len(find_word.split()) == 1:
        # check matches in normal image (black text)
        matching_indices = [i for i, word in enumerate(data['text']) if find_word.lower() in word.lower()]
        matches = [
            {
                'left':data['left'][i],
                'top':data['top'][i],
                'width':data['width'][i],
                'height':data['height'][i]
            }

            for i in matching_indices
            ]
        matches = sorted(matches, key=lambda x:x[sort_key], reverse=reversed_sort)

        # check matches in inverted image (white text)
        matching_indices = [i for i, word in enumerate(data_i['text']) if find_word.lower() in word.lower()]
        matches_i = [
            {
                'left':data_i['left'][i],
                'top':data_i['top'][i],
                'width':data_i['width'][i],
                'height':data_i['height'][i]
            }

            for i in matching_indices
            ]
        matches_i = sorted(matches_i, key=lambda x:x[sort_key], reverse=reversed_sort)

        if len(matches) == expected_match_num:
            match_loc = matches[0]
        elif len(matches_i) == expected_match_num:
            match_loc = matches_i[0]
        else:
            raise ValueError(f"Something went wrong, not on the expected screen to select {find_word}")
    else:
        match_loc = find_phrase(find_word, data)
        if len(match_loc) ==0:
            match_loc = find_phrase(find_word, data_i)
            

    if click_local == 'center':
        pyautogui.click(x=(match_loc['left']+match_loc['width']//2), y=(match_loc['top']+match_loc['height']//2), button=click_type)


def match_image_regardless_of_sz(im_name, screenshots_dir, return_closest_center=True, return_closest_point=False, closest_coordsx=0, closest_coordsy=0):
    img1 = cv2.imread(os.path.join(screenshots_dir, im_name),cv2.IMREAD_GRAYSCALE)          # queryImage
    # plt.imshow(img1), plt.show()

    screenshot = pyautogui.screenshot()
    scrnshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    img2 = cv2.cvtColor(scrnshot, cv2.COLOR_BGR2GRAY)

    # Initiate SIFT detector
    sift = cv2.SIFT_create()
    
    # find the keypoints and descriptors with SIFT
    kp1, des1 = sift.detectAndCompute(img1,None)
    kp2, des2 = sift.detectAndCompute(img2,None)

    if des1 is None or des2 is None:
        raise ValueError("no descriptors")
    
    # BFMatcher with default params
    bf = cv2.BFMatcher(cv2.NORM_L2, crossCheck=True)
    matches = bf.match(des1,des2)
    
    # # # Sort them in the order of their distance.
    matches = sorted(matches, key = lambda x:x.distance)
    
    if return_closest_center:
        # Center of image
        center_x = img2.shape[0] //2
        center_y = img2.shape[1] //2
        center_point = (center_x, center_y)

    if return_closest_point:
        center_point = (closest_coordsx, closest_coordsy)
    
    if return_closest_center or return_closest_point:
        # find closest match
        closest_match = None
        min_distance = float('inf')
        save_keypoint = 0
        for match in matches:
            keypoint2 = kp2[match.trainIdx]
            keypoint2_pt = (int(keypoint2.pt[0]), int(keypoint2.pt[1]))

            # calculate distance from center
            distance = np.linalg.norm(np.array(keypoint2_pt) - np.array(center_point))
            if distance < min_distance:
                min_distance = distance
                closest_match = match
                save_keypoint = keypoint2_pt
        
    # # # # Draw first 10 matches.
    # img3 = cv2.drawMatches(img1,kp1,img2,kp2,matches[:10],None,flags=cv2.DrawMatchesFlags_NOT_DRAW_SINGLE_POINTS)
    
    # plt.imshow(img3),plt.show()
    if return_closest_center:
        return(save_keypoint)
    else:
        return(matches)

def write_to_input_field(text):
    pyperclip.copy(str(text))
    pyautogui.hotkey('ctrl', 'a')   
    pyautogui.hotkey('ctrl', 'v')


# take screenshot, move down, take screenshot, find difference in highlighted areas, center of the two and select the bottom one
# as started at the top (take into account if you would need to move up what changes would need to be made)
def move_to_location_higlighted(name, categories):
        time.sleep(1) # added to ensure it has a moment of the screen adjusting before starting to hunt for differences
        
        screenshot1, _ = take_screenshot(ialpha=7, ibeta=0, alpha=1.5, beta=0, show_screenshot=False)
        pyautogui.press('down', presses=categories.index(name)) 
        screenshot2, _ = take_screenshot(ialpha=7, ibeta=0, alpha=1.5, beta=0, show_screenshot=False)
        difference = cv2.absdiff(screenshot1, screenshot2)  
        contours, _ = cv2.findContours(difference, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        screenshot1 = cv2.cvtColor(screenshot1, cv2.COLOR_GRAY2BGR)

        centers = []
        center = (0,0)
        for contour in contours:
            if cv2.contourArea(contour) > 100:
                M = cv2.moments(contour)
                if M["m00"] !=0:
                    cx = int(M["m10"] /M["m00"])
                    cy = int(M["m01"] /M["m00"])
                    if cy > center[1]:
                        center = (cx, cy)
                    
                    
                    centers.append((cx, cy))
                    cv2.circle(screenshot1, (cx, cy), 25, (0,255,0), -1)

        pyautogui.moveTo(center[0], center[1])
        return center

def generate_tecan_files(df, output_folder):
    # need to press alt or doesnt work for some reason
    pyautogui.press("alt")
    w = WindowMgr()
    w.find_window_wildcard("FluentControl")
    w.set_foreground() 
    click_word_on_screen('Labware', expected_match_num=1)
    click_word_on_screen('microplate', expected_match_num=1)
    for i, row in df.iterrows():
        well_bottom_shape = row["well_bottom_shape"]
        plate_x = row["plate_length_mm"]
        plate_y = row["plate_width_mm"] 
        plate_z = row["plate_height_mm"]    
        plate_name = clean_string(row["plate_name"]) # Name does not accept ",-=!?" etc, change this later  to keep spaces
        plate_rows = row["num_rows"]
        plate_cols = row["num_columns"]

        # check all are non negative
        # Z-travel = plate height + 10mm
        # z-start = plate height
        # z-dispense =  0.8(max well depth) + z-bottom #20% of depth
        # z-max = (plate height - max well depth) + 0.5mm
        # z-bottom = plate height - max well depth 
        z_travel = plate_z + 10
        z_start = plate_z
        z_bottom = plate_z - row["well_max_depth_mm"]
        z_dispense = 0.8 *  row["well_max_depth_mm"] + z_bottom
        z_max = z_bottom + 0.5 
        
        # for determining where the plate will be located on the GUI
        existing_categories = ["1536 Well", "24 Well Flat", "384 Well", "384 Well LowVol HiBase", "384 Well LowVol LoBase", "384 Well LowVol Round Corning", "384 Well PCR", "48 Well Flat", "96 Well Crystal Quick", "96 Well Flat", "96 Well Round", "96 Well Skirted PCR", "96 Well TeVacs Filterplate", "96 Well TeVacs Filterplate High", "PCR Adapter 96 Well"]

        a1_left_offset = row["offset_left_to_a1_mm"] + (row["well_length_mm"] / 2)
        a1_bottom_offset = row["offset_left_to_a1_mm"] + (row["well_width_mm"] / 2) # assumes to center of well and actually measures from bottom (so bottom edge is 0) 
        h12_left_offset =  row["plate_length_mm"] - row["offset_right_to_final_well_mm"] - (row["well_length_mm"] / 2) # is the distance from the left of the plate to the center of bottom right most well 
        h12_bottom_offset = row["offset_bottom_to_final_well_mm"] + (row["well_width_mm"] / 2) # is the distance from the bottom of the plate to the center of bottom right most well  # may be wrong
        cone_compartment_height = row["well_max_depth_mm"]  - row["bottom_start_depth_mm"]
        cone_compartment_top_diameter = row["well_width_mm"]
        sphere_compartment_diameter =  row["well_width_mm"]
        sphere_compartment_height = sphere_compartment_diameter / 2 
        cuboid_compartment_length = row["well_length_mm"]
        cuboid_compartment_width = row["well_width_mm"]
        cuboid_compartment_height = row["well_length_mm"]  

        click_word_on_screen('96 well flat', expected_match_num=1, click_type='right')
        current_position = pyautogui.position()
        pyautogui.press('down', presses=2)
        pyautogui.press('enter')
        # click_word_on_screen('Duplicate 96 Well Flat', expected_match_num=1) # could change to do this with
                            
        # 30 second delay   
        time.sleep(10)

        # drag to worktable
        #######################################
        # take screenshot
        scrnshot, screenshot_inv = take_screenshot()

        # convert to text data      
        data = pytesseract.image_to_data(scrnshot, output_type=pytesseract.Output.DICT)
        data_i = pytesseract.image_to_data(screenshot_inv, output_type=pytesseract.Output.DICT)

        pyautogui.hotkey('alt', 'w')
        pyautogui.press('down', presses=2)
        pyautogui.press('enter')                                
        pyautogui.press('tab')

        r_center = move_to_location_higlighted("96 Well Round", existing_categories) # go here and not "96 Well Flat_1" because new plate was added and this location is where it is and no need to add it to existing_categories
        gap = pyautogui.position()[1] - current_position[1] 
        # cv2.imshow('Difference', screenshot1)
        # cv2.waitKey(5000)
        # cv2.destroyAllWindows()

        # give it time to move
        time.sleep(2)

        match = match_image_regardless_of_sz('carrier part larger.PNG', screenshots_dir)
        # -15 as match is to right edge and want to be more central
        carrier_positon = (match[0]-15, match[1])
        pyautogui.dragTo(carrier_positon[0], carrier_positon[1], 2, button='left')

        # select plate so can edit things about it
        pyautogui.click()

        # click on general tab
        scrnshot, screenshot_inv = take_screenshot()
        data_i = pytesseract.image_to_data(screenshot_inv, output_type=pytesseract.Output.DICT)
        coords = find_phrase("General Settings", data_i)
        pyautogui.moveTo(coords['left'], coords['top'])
        pyautogui.click()
        pyautogui.press('tab')
        pyautogui.press('space')
        pyautogui.press('tab', presses=9)

        # rename plate
        write_to_input_field(plate_name)

        # reduce tab again
        pyautogui.press('tab', presses=9)
        pyautogui.press('space')

        # open next tab
        pyautogui.press('tab')
        pyautogui.press('space')

        # change X
        pyautogui.press('tab', presses=8)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.write(str(plate_x))

        # change y
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.write(str(plate_y))  

        # change Z
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.write(str(plate_z))  

        # change number of columns
        pyautogui.press('tab', presses=2)
        write_to_input_field(plate_cols)

        # change number of rows
        pyautogui.press('tab', presses=3)
        write_to_input_field(plate_rows)

        match = match_image_regardless_of_sz('carrier part larger.PNG', screenshots_dir)

        # double click on the carrier so that the well positions can be selected
        carrier_positon = (match[0]-30, match[1])
        pyautogui.moveTo(carrier_positon)
        pyautogui.click(clicks=2)

        # go back to the general settings and move to the well position input fields
        scrnshot, screenshot_inv = take_screenshot()
        data_i = pytesseract.image_to_data(screenshot_inv, output_type=pytesseract.Output.DICT)
        coords = find_phrase("currently", data_i)
        pyautogui.moveTo(coords['left'], coords['top'])
        pyautogui.click()
        pyautogui.press('tab', presses=15)

        # fill in coords for a1 left
        write_to_input_field(a1_left_offset)
        pyautogui.press('tab', presses=1)

        # fill in val for a1 bottom
        write_to_input_field(a1_bottom_offset)

        # unlink val
        pyautogui.press('tab', presses=1)
        pyautogui.press('space')

        # move to h12 left
        pyautogui.press('tab', presses=3)

        # fill in val for h12 left
        write_to_input_field(h12_left_offset)
        pyautogui.press('tab', presses=1)

        # fill in val for h12 bottom
        write_to_input_field(h12_bottom_offset)

        # unlink z things
        pyautogui.press('tab', presses=14)
        pyautogui.press('space')

        # move back to z travel
        for i in range(0, 7):
            pyautogui.hotkey('shift', 'tab') # to go backwards

        # set z vals
        for z_val in [z_travel, z_start, z_dispense, z_max, z_bottom]:
            pyperclip.copy(str(z_val))
            pyautogui.hotkey('ctrl', 'v')
            if (z_val == z_dispense):   
                pyautogui.press('tab', presses=4)
            else:
                pyautogui.press('tab', presses=3)

        # collapse position section
        pyautogui.press('tab', presses=3)
        pyautogui.press('space')

        # expand compartment defintion
        pyautogui.press('tab')
        pyautogui.press('space')

        # select shape option
        pyautogui.press('tab', presses=7)

        # press down arrow key till you get to the right shape

        # only partial sphere, cone, cuboid is used
        map_arrow_to_shape = {
            "Partial Sphere": 1,
            "Cylinder": 2,
            "Cone": 3,
            "Cuboid": 4,
            "Pyramid": 5,
            "Trapezoid": 6,
            "User Defined": 7
        }   

        if (well_bottom_shape == "U-Bottom"):
            # Select partial sphere
            pyautogui.press('down', presses=map_arrow_to_shape["Partial Sphere"])

            # Press add button 
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter', presses=1)

            # Move to first input field and input values (height)
            pyautogui.press('tab', presses=4)
            write_to_input_field(sphere_compartment_height)

            pyautogui.press('tab')

            # Second input field (diameter)
            write_to_input_field(sphere_compartment_diameter)
        elif (well_bottom_shape == "V-Bottom"):
            # Select Cone
            pyautogui.press('down', presses=map_arrow_to_shape["Cone"])

            # Press add button 
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter', presses=1)

            # Move to first input field and input values (height)
            pyautogui.press('tab', presses=4)
            write_to_input_field(cone_compartment_height)
            pyautogui.press('tab')

            # Second input field (bottom_diameter is always 0)
            write_to_input_field(0)

            pyautogui.press('tab')

            # Third input field (top_diameter)
            write_to_input_field(cone_compartment_top_diameter)

        else:
            # Select cuboid
            pyautogui.press('down', presses=map_arrow_to_shape["Cuboid"])
            
            # Press add button
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter', presses=1)
            pyautogui.press('tab', presses=4)

            # Move to first input field and input values (height)
            write_to_input_field(cuboid_compartment_height)

            pyautogui.press('tab')

            # Move to second input field and input values (length)
            write_to_input_field(cuboid_compartment_length)
            pyautogui.press('tab')

            # Move to third input field and input values (width)
            write_to_input_field(cuboid_compartment_width)

        # delete existing shape
        pyautogui.press('tab', presses=3)
        pyautogui.press('enter')

        # close compartment section
        pyautogui.press('tab', presses=10)
        pyautogui.press('space')

        # go save and close 
        pyautogui.press('tab', presses=8)
        pyautogui.press('space')

        # Export selection
        pyautogui.hotkey('alt', 'd')
        pyautogui.press('enter')

        # Select labware
        pyautogui.press('tab')
        pyautogui.press('down', presses=7)
        pyautogui.press('enter')

        # Select microplate
        pyautogui.press('down', presses=8)
        pyautogui.press('enter')

        ### Select plate that was just created
        # Go down to that created plate, alphabetically sorted
        existing_categories.append(plate_name)
        existing_categories.sort()
        pyautogui.press('down', presses=existing_categories.index(plate_name) + 1) 

        # Select add option (add with  dependencies right now?)
        pyautogui.press('tab', presses=2)
        pyautogui.press('space')

        # Export button
        pyautogui.press('tab', presses=5)
        pyautogui.press('space')

        # export window, export to get file with plate name
        pyperclip.copy(plate_name)
        pyautogui.hotkey('ctrl', 'v')

        # go change the directory to save into results folder
        # NUMBER OF TABS MAY VARY FROM PC TO PC so change it?
        pyautogui.press('tab', presses=6)
        pyperclip.copy(output_folder)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        # go from directory field to save button
        # NUMBER OF TABS MAY VARY FROM PC TO PC so change it?
        pyautogui.press('tab', presses=9)
        pyautogui.press('enter')

        # file finished exporting, alt w to open window, down to select "controlbar", tab to select the plates, go all the way up to highlight the first plate
        pyautogui.hotkey('alt', 'w')
        pyautogui.press('down', presses=2)
        pyautogui.press('enter')
        pyautogui.press('tab', presses=7)
        pyautogui.press('up', presses=len(existing_categories))

        # find location of the plate in the controlbar, doesnt work if the plate  is located at the top 
        if existing_categories.index(plate_name) == 0:
            temp = move_to_location_higlighted("24 Well Flat", existing_categories)
            center = (temp[0], temp[1] - gap * 2)
        else:
            center = move_to_location_higlighted(plate_name, existing_categories)

        # give it time to move
        time.sleep(2)

        # mouse now hovering the added plate, right click and remove it
        pyautogui.click(center[0], center[1], button='right')
        pyautogui.press('down', presses=3)
        pyautogui.press('enter')

        # give it time to delete
        time.sleep(10)

        # After deleting and waiting out the delay, go back to highlighting the first plate (white highlight confuses the OCR, so just make it consistent)
        pyautogui.press('up', presses=len(existing_categories))

        # remove the plate, move on to next plate
        existing_categories.remove(plate_name)
        