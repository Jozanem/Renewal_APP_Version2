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
        f"POLICY RENEWAL: {Policy_Number} – DOMESTIC PROTECTOR POLICY\n"
    )
    set_font_style(first_paragraph, bold=True, underline=True, color=RGBColor(0, 0, 0))

    # Create content_paragraph1 with the renewal date in bold
    content_paragraph1 = doc.add_paragraph()
    content_paragraph1.add_run(
        "The above-mentioned policy falls due for renewal on the "
    ).font.name = "Cambria"
    renewal_date_run = content_paragraph1.add_run(str(cover_start_date))
    renewal_date_run.font.name = "Cambria"  # Ensure Cambria font
    renewal_date_run.font.size = Pt(11)  # Set size to 11
    renewal_date_run.bold = True  # Make only renewal_date bold

    # Create content_paragraph2 with the "2024/2025" in bold
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

    Content_paragraph3 = doc.add_heading("HOUSEOWNERS/HOUSEHOLDERS SECTION")
    set_font_style(Content_paragraph3, bold=True, underline=True)

    # Sub-header
    Content_paragraph4 = doc.add_heading(
        "Cover is against the following perils:", level=2
    )
    set_font_style(
        Content_paragraph4, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    # Add the bullet points
    bullet_points = [
        "Fire, lightning & explosion",
        "Storm, wind, water, hail, or snow excluding loss or damage caused by subsistence or landslip and loss or damage to gates, retaining walls and all types of fences.",
        "Earthquake",
        "Bursting or overflow of water tanks, apparatus or pipes including damage to the apparatus or pipes.",
        "Accidental impact to the exterior portion of the building",
        "Theft or attempted theft accompanied by forcible and violent entry or exit Riot, civil commotion, strikes, labour disturbances and persons of malicious intent.",
    ]

    for point in bullet_points:
        bullet_paragraph1 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph1)

    Content_paragraph5 = doc.add_paragraph(
        "This policy has been further enriched with the following extensions aimed at providing adequate insurance cover to you:"
    )
    set_font_style(
        Content_paragraph5,
        bold=True,
    )

    # Add the bullet points
    bullet_points = [
        "Capital improvements; additional improvements on existing structures are automatically covered up to 15% of the sum insured. However, as soon as reasonably possible, you should inform the Corporation of those improvements.",
        "Loss of rent; if the buildings are lent out to a tenant, the client can claim up to 25% of the sum insured in respect of loss of rent whilst the buildings are uninhabitable.",
    ]

    for point in bullet_points:
        bullet_paragraph2 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph2)

    Content_paragraph6 = doc.add_heading("Policy enhancement", level=2)
    set_font_style(Content_paragraph6, bold=True, underline=True)

    Content_paragraph7 = doc.add_paragraph(
        "Our policy has been enhanced to include the following:"
    )

    Content_paragraph8 = doc.add_paragraph()
    # Trauma
    part1_run = Content_paragraph8.add_run(
        "Trauma counseling cover, up to a limit of E2, 500."
    )
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # trauma
    part2_run = Content_paragraph8.add_run(
        "per incident. The incident itself must be caused by an insured peril. The counseling must be conducted by a registered practicing psychologist.\n"
    )
    part2_run.font.name = "Cambria"
    part2_run.font.size = Pt(11)

    # Emergency
    part3_run = Content_paragraph8.add_run(
        "Emergency alternative accommodation cover, up to a limit of E5, 000.00 "
    )
    part3_run.font.name = "Cambria"
    part3_run.font.size = Pt(11)
    part3_run.bold = True

    # Emergency
    part4_run = Content_paragraph8.add_run(
        "in the event the premises are damaged by any of the insured perils and rendered uninhabitable.\n"
    )
    part4_run.font.name = "Cambria"
    part4_run.font.size = Pt(11)

    # Power Surge
    part5_run = Content_paragraph8.add_run("Power Surge Cover (free)")
    part5_run.font.name = "Cambria"
    part5_run.font.size = Pt(11)
    part5_run.bold = True

    # Power Surge
    part6_run = Content_paragraph8.add_run(
        "– Cover is extended to include power surge damages for a limit of E20, 000.00.\n"
    )
    part6_run.font.name = "Cambria"
    part6_run.font.size = Pt(11)

    # Geyser
    part7_run = Content_paragraph8.add_run("Geyser cover extension\n")
    part7_run.font.name = "Cambria"
    part7_run.font.size = Pt(11)
    part7_run.bold = True

    # Geyser
    part8_run = Content_paragraph8.add_run(
        " Geyser is covered against accidental damage and bursting subject to the following provisos;\n"
        "-	Only one claim per location in three years is allowed;\n"
        "-	An excess of 10% of claim minimum E750.00 will apply;\n"
        "-	In the event that there is no drip tray, an excess of 10% of claim minimum E750.00 will apply in respect of water damage ;\n"
        "-	Only geysers aged up to 10 years are covered on this policy."
    )
    part8_run.font.name = "Cambria"
    part8_run.font.size = Pt(11)

    Content_paragraph9 = doc.add_paragraph(
        "NB: We advise you to ensure that the sum insured for your property(s) represents the current market value and we strongly recommend that you consult with a professional property evaluator to get the correct replacement value of your property to avoid disputes at claims stage."
    )
    set_font_style(Content_paragraph9, bold=True)

    Content_paragraph10 = doc.add_paragraph()
    # Contact
    part1_run = Content_paragraph10.add_run(
        "Kindly provide us with current contact details e.g. Cell Number, E-mail Address etc.\n"
    )
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # trauma
    part2_run = Content_paragraph10.add_run(
        "Please advise our office on any amendments required regarding the basis of cover, we will gladly assist you."
    )
    part2_run.font.name = "Cambria"
    part2_run.font.size = Pt(11)

    Content_paragraph11 = doc.add_paragraph("Warm Regards\n")
    set_font_style(Content_paragraph11, bold=True)

    Content_paragraph12 = doc.add_paragraph()

    # Department
    part1_run = Content_paragraph12.add_run("UNDERWRITING TEAM\n")
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # email
    part2_run = Content_paragraph12.add_run("e-mail: underwriting@sric.sz")
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
