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
        f"POLICY RENEWAL: {Policy_Number} – MOTOR SPECIAL FLEET"
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
    content_paragraph2 = doc.add_paragraph()

    # Regular text before the bold part
    part1_run = content_paragraph2.add_run(
        "Please find attached the renewal notice for the "
    )
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)

    # The bold "2025/2026"
    year_run = content_paragraph2.add_run("2026/2027")
    year_run.font.name = "Cambria"
    year_run.font.size = Pt(11)
    year_run.bold = True  # Make "2025/2026" bold

    # Regular text after the bold part
    part2_run = content_paragraph2.add_run(
        " period of insurance. It is important to note that no cover will be afforded and no claim will be paid unless premiums have been paid on or before renewal."
    )
    part2_run.font.name = "Cambria"
    part2_run.font.size = Pt(11)

    content_paragraph3 = doc.add_paragraph(
        "Please note that cover is against any accidental physical damage to the vehicle, except what is explicitly excluded such as unroadworthy vehicle, being driven outside the territorial limits, unlicensed drivers and being driven by someone under the influence of alcohol. Covered perils include; collision, overturning, theft or attempted theft, straying off the road and hitting any object, hijack, fire, lightning and explosion."
    )

    Content_paragraph4 = doc.add_paragraph(
        "NB: We advise you to ensure that the sum insured for your vehicle(s) represents the correct market value and we strongly recommend that you consult with your car Dealer to get the correct value."
    )
    set_font_style(Content_paragraph4, bold=True)

    Content_paragraph5 = doc.add_paragraph()
    # Contact
    part1_run = Content_paragraph5.add_run(
        "Kindly provide us with current contact details e.g. Cell Number, E-mail Address etc.\n"
    )
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # trauma
    part2_run = Content_paragraph5.add_run(
        "Please advise our office on any amendments required regarding the basis of cover, we will gladly assist you."
    )
    part2_run.font.name = "Cambria"
    part2_run.font.size = Pt(11)

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
