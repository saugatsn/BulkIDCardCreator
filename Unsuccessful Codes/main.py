from PIL import Image, ImageDraw, ImageFont
import openpyxl

# Load template image
template = Image.open('id_card_template.png')

# Load Excel file
wb = openpyxl.load_workbook('student_data.xlsx')
ws = wb.active

# Set up font styles
name_font = ImageFont.truetype('arial.ttf', 48)
info_font = ImageFont.truetype('arial.ttf', 24)
valid_till_font = ImageFont.truetype('arial.ttf', 24)

# Loop through each row in the Excel file
for row in ws.iter_rows(min_row=2):
    # Get data from the row
    name, crn, level, program, cell, dob, blood_group, citizenship, email, valid_till = [cell.value for cell in row]
    
    # Load photo with the same name as student's name
    photo = Image.open(f'{name}.jpg')

    # Resize photo to fit in the ID card
    photo = photo.resize((200, 200))

    # Paste photo onto template at (x, y) position
    template.paste(photo, (50, 100))

    # Add text to the ID card
    draw = ImageDraw.Draw(template)
    draw.text((300, 110), name, font=name_font, fill='black')
    draw.text((300, 170), f"CRN: {crn}", font=info_font, fill='black')
    draw.text((300, 200), f"Level: {level}", font=info_font, fill='black')
    draw.text((300, 230), f"Program: {program}", font=info_font, fill='black')
    draw.text((300, 260), f"Cell: {cell}", font=info_font, fill='black')
    draw.text((300, 290), f"DOB: {dob}", font=info_font, fill='black')
    draw.text((300, 320), f"Blood Group: {blood_group}", font=info_font, fill='black')
    draw.text((300, 350), f"Citizenship: {citizenship}", font=info_font, fill='black')
    draw.text((300, 380), f"Email: {email}", font=info_font, fill='black')
    draw.text((700, 160), f"Valid Till:\n{valid_till}", font=valid_till_font, fill='black', align='center')
    
    # Save the ID card as JPEG file
    template.save(f'{name}_id_card.jpg')
    
    # Reset template for next student
    template = Image.open('id_card_template.png')
