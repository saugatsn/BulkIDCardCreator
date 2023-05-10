from PIL import Image, ImageDraw, ImageFont
import openpyxl
import os

# Define the path to the ID card template image and the Excel file
template_path = "id_card_template.png"
excel_file_path = "student_data.xlsx"

# Define the font paths for different font faces
arial_path = "arial.ttf"
times_new_roman_path = "times.ttf"
calibri_path = "calibri.ttf"

# Define the font sizes for different fields
name_size = 48
crn_size = 24
level_size = 24
program_size = 24
cell_size = 24
dob_size = 24
blood_group_size = 24
citizenship_size = 24
email_size = 24
valid_till_size = 16

# Define the font colors for different fields
name_color = (0, 0, 0)  # black
crn_color = (255, 255, 255)  # white
level_color = (255, 255, 255)  # white
program_color = (255, 255, 255)  # white
cell_color = (255, 255, 255)  # white
dob_color = (255, 255, 255)  # white
blood_group_color = (255, 255, 255)  # white
citizenship_color = (255, 255, 255)  # white
email_color = (255, 255, 255)  # white
valid_till_color = (0, 0, 0)  # black

# Define the font faces for different fields
name_font = ImageFont.truetype(arial_path, name_size)
crn_font = ImageFont.truetype(calibri_path, crn_size)
level_font = ImageFont.truetype(calibri_path, level_size)
program_font = ImageFont.truetype(calibri_path, program_size)
cell_font = ImageFont.truetype(calibri_path, cell_size)
dob_font = ImageFont.truetype(calibri_path, dob_size)
blood_group_font = ImageFont.truetype(calibri_path, blood_group_size)
citizenship_font = ImageFont.truetype(calibri_path, citizenship_size)
email_font = ImageFont.truetype(calibri_path, email_size)
valid_till_font = ImageFont.truetype(times_new_roman_path, valid_till_size)

# Define the positions for different fields
name_pos = (300, 350)
crn_pos = (200, 420)
level_pos = (200, 470)
program_pos = (200, 520)
cell_pos = (200, 570)
dob_pos = (200, 620)
blood_group_pos = (200, 670)
citizenship_pos = (200, 720)
email_pos = (200, 770)
valid_till_pos = (825, 610)

# Define the angle for the ValidTill field
valid_till_angle = -90

# Define the radius for the circular photo
photo_radius = 100

# Load the ID card template image
id_card = Image.open(template_path)

# Load the data from the Excel file
wb = openpyxl.load_workbook(excel_file_path)
sheet = wb.active
rows = sheet.iter_rows(min_row=2, values_only=True)

# Loop through each row of data in the Excel file
for i, row in enumerate(rows):
    # Extract the data from the row
    name, crn, level, program, cell, dob, blood_group, citizenship, email, valid_till = row
    
    # Load the photo for this student
    photo_path = os.path.join("photos", name + ".jpg")
    try:
        photo = Image.open(photo_path)
    except FileNotFoundError:
        # If the photo file is not found, skip this row
        print(f"Warning: photo file not found for student {name}")
        continue
    
    # Resize the photo to a circle and paste it onto the ID card template
    photo = photo.resize((2*photo_radius, 2*photo_radius))
    mask = Image.new("L", photo.size, 0)
    draw_mask = ImageDraw.Draw(mask)
    draw_mask.ellipse((0, 0, 2*photo_radius, 2*photo_radius), fill=255)
    photo.putalpha(mask)
    id_card.paste(photo, (100, 300), photo)
    
    # Draw the text onto the ID card template
    draw = ImageDraw.Draw(id_card)
    draw.text(name_pos, name, font=name_font, fill=name_color)
    draw.text(crn_pos, "CRN: " + str(crn), font=crn_font, fill=crn_color)
    draw.text(level_pos, "Level: " + str(level), font=level_font, fill=level_color)
    draw.text(program_pos, "Program: " + str(program), font=program_font, fill=program_color)
    draw.text(cell_pos, "Cell: " + str(cell), font=cell_font, fill=cell_color)
    draw.text(dob_pos, "DOB: " + str(dob), font=dob_font, fill=dob_color)
    draw.text(blood_group_pos, "Blood Group: " + str(blood_group), font=blood_group_font, fill=blood_group_color)
    draw.text(citizenship_pos, "Citizenship: " + str(citizenship), font=citizenship_font, fill=citizenship_color)
    draw.text(email_pos, "Email: " + str(email), font=email_font, fill=email_color)
    draw.text(valid_till_pos, "Valid Till: " + str(valid_till), font=valid_till_font, fill=valid_till_color)
    draw.text(valid_till_pos, "Valid Till: " + str(valid_till), font=valid_till_font, fill=valid_till_color, anchor="ra", angle=valid_till_angle)
    
    # Save the ID card as a JPG file
    id_card.save(f"id_card_{i+1}.jpg")
    
    # Reset the ID card template for the next student
    id_card = Image.open(template_path)
