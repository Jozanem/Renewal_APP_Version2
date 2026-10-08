from services.letter_styles import set_font_style, add_header_footer_first_page
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import re


def generate(
    client_name,
    address1,
    address2,
    Policy_Number,
    cover_start_date,
    output_folder,
    header_image_path,
    footer_image_path,
):
    # Create a new Document object
    doc = Document()

    # Call the function to add header and footer to the first page only
    add_header_footer_first_page(doc, header_image_path, footer_image_path)

    # Add date, align to right, bold, and set font
    today_date = datetime.today().strftime("%d %B %Y")  # Get today's date
    date_paragraph = doc.add_paragraph(today_date)
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font_style(date_paragraph, bold=True)

    # Add client name and address, left-aligned, bold, and set font
    client_name_paragraph = doc.add_paragraph(str(client_name))
    address1_paragraph = doc.add_paragraph(str(address1))
    address2_paragraph = doc.add_paragraph(str(address2))

    for paragraph in [client_name_paragraph, address1_paragraph, address2_paragraph]:
        set_font_style(paragraph, bold=True)

    # Add greeting
    greeting_paragraph = doc.add_paragraph("Dear Sir/Madam,")
    set_font_style(greeting_paragraph)

    # Add the first paragraph, bold, underlined, and set font
    first_paragraph = doc.add_heading(
        f"POLICY RENEWAL: {Policy_Number} – WORKMEN'S COMPENSATION"
    )
    set_font_style(first_paragraph, bold=True, underline=True)

    # Create content_paragraph1 with the renewal date in bold
    content_paragraph1 = doc.add_paragraph()
    content_paragraph1.add_run(
        "The above-mentioned policy falls due for renewal on the "
    ).font.name = "Cambria"
    renewal_date_run = content_paragraph1.add_run(str(cover_start_date))
    renewal_date_run.font.name = "Cambria"  # Ensure Cambria font
    renewal_date_run.font.size = Pt(11)  # Set size to 11
    renewal_date_run.bold = True  # Make only renewal_date bold

    # Create content_paragraph2 with the "2025/2026" in bold
    Content_paragraph2 = doc.add_paragraph(
        "Kindly note that the premium on the renewal notice is provisional subject to submission of the attached wages declaration form."
    )
    set_font_style(Content_paragraph2, bold=True)

    Content_paragraph3 = doc.add_paragraph(
        "It is important to note that no cover will be afforded and no claim will be paid, unless premiums have been paid on or before renewal."
    )
    set_font_style(Content_paragraph3, bold=True)

    Content_paragraph4 = doc.add_heading("Scope of cover", level=2)
    set_font_style(Content_paragraph4, bold=True, color=RGBColor(0, 0, 0))

    bullet_points = [
        "Cover for employees against occupational or employment injuries.",
        "It is a statutory requirement for every organization in Eswatini to have this policy.",
        "Cover is not limited within the organization’s premises. It starts immediately the employee leaves his/her place of residence, drives to/from the place of employment or boards a public transport, goes to buy lunch. During all this time, the employee is covered by the workmen’s compensation insurance.",
    ]

    for point in bullet_points:
        bullet_paragraph1 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph1)

    Content_paragraph5 = doc.add_heading("The benefits and payments are as follows;")
    set_font_style(Content_paragraph5, bold=True, color=RGBColor(0, 0, 0))

    bullet_points = [
        "Death – the employee will be paid 48 times his monthly salary.",
        "Permanent disability – the employee will be paid 52 times his monthly salary at a certain multiplier which represents the severity of the injuries. The multiplier is allocated by a professional medical practitioner.",
        "Total temporal disablement.",
        "Medical expenses – The insurer pays up to E200,000.00per person per accident. In this regard, we emphasize that we only refund the medical bills.",
        "Funeral cover of E20,000.00 per person in the event of death caused by occupational injuries",
        "Employer liability of E1,000,000.00 in the aggregate ",
    ]

    for point in bullet_points:
        bullet_paragraph2 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph2)

    Content_paragraph5 = doc.add_paragraph(
        "Should you need any assistance on this policy, do not hesitate to call or to come to our offices we will gladly assist you."
    )

    Content_paragraph6 = doc.add_paragraph("Warm Regards\n")
    set_font_style(Content_paragraph6, bold=True)

    Content_paragraph7 = doc.add_paragraph()

    # Department
    part1_run = Content_paragraph7.add_run("UNDERWRITING TEAM\n")
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # email
    part2_run = Content_paragraph7.add_run("e-mail: underwriting@sric.sz")
    part2_run.font.name = "Italic"
    part2_run.font.size = Pt(11)
    part2_run.bold = True

    # --------------------------------------------------
    # SAVE DOCUMENT
    # --------------------------------------------------

    # SAFE POLICY NUMBER
    policy_number = str(Policy_Number).strip()

    if not policy_number:
        policy_number = "UNKNOWN"

    # SANITIZE (Windows safe filename)
    policy_number = re.sub(r'[\\/*?:"<>|]', "_", policy_number)

    # CREATE OUTPUT FOLDER
    os.makedirs(output_folder, exist_ok=True)

    # OUTPUT PATH
    output_path = os.path.join(output_folder, f"{policy_number}.docx")

    # SAVE FILE
    doc.save(output_path)

    return output_path
