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
    first_paragraph = doc.add_paragraph(
        f"POLICY RENEWAL: {Policy_Number} – DOMESTIC PROTECTOR POLICY"
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

    # The bold "2024/2025"
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

    content_paragraph3 = doc.add_heading("MOTOR SECTION – COMPREHENSIVE")
    set_font_style(
        content_paragraph3, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    Content_paragraph4 = doc.add_paragraph(
        "Please note that cover is against any accidental physical damage to the vehicle, except what is explicitly excluded such as unroadworthy vehicle, being driven outside the territorial limits, unlicensed driver and being driven by someone under the influence of alcohol. Covered perils include; collision, overturning, theft or attempted theft, straying off the road and hitting any object, hijack, fire, lightning and explosion."
    )

    Content_paragraph5 = doc.add_paragraph(
        "The policy has an in-built third-party liability of E2, 000, 000.00 to protect you from lawsuits arising from injuries to pedestrians and damages to third party properties."
    )

    Content_paragraph6 = doc.add_paragraph(
        "Following an increased demand in the market for all motor vehicle classes in general, we have noted an upward trend in prices. Hence some vehicles are appreciating in value instead of depreciating. Consequently, we advise you to monitor this trend and revise sums insured for your vehicle accordingly to ensure that the vehicle is adequately insured."
    )
    set_font_style(Content_paragraph6, bold=True)

    Content_paragraph7 = doc.add_paragraph(
        "We advise you to ensure that the sum insured for your property(s) represents the current market value and we strongly recommend that you consult with your car dealer to get the correct value of your vehicles to avoid disputes at claims stage."
    )
    set_font_style(Content_paragraph7, bold=True)

    Content_paragraph8 = doc.add_heading("Average Clause", level=2)
    set_font_style(
        Content_paragraph8, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    Content_paragraph9 = doc.add_paragraph(
        "We wish to draw your attention to the application of average clause where the insurer discovers that the sum insured of the insured vehicle is less than the actual market value of vehicle at the time of loss or damage. In such instances you shall be considered self-insured for the difference and shall bear a ratable share of the loss accordingly. Every vehicle, if more than one, shall be separately subjected to this condition. The onus to ensure correct valuation rests on the insured."
    )
    set_font_style(Content_paragraph9, bold=True)

    Content_paragraph10 = doc.add_heading("Policy Enhancements", level=2)
    set_font_style(
        Content_paragraph10, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    Content_paragraph11 = doc.add_paragraph(
        "Our comprehensive motor insurance cover has been further enhanced to include for free the following key extensions;"
    )

    # Add the bullet points
    bullet_points = [
        "Free emergency accommodation limited to E1000.00 per claim following accidental damage rendering vehicle not drivable, theft or hijack of the insured vehicle;",
        "Free funeral cover limited to E5000.00 per claim following the death of the Insured in the insured vehicle consequent a motor vehicle accident;",
        "Free emergency rental cover limited to E1500.00 following accidental damage rendering vehicle not drivable, theft or hijack of the insured vehicle;",
        "Free trauma counseling limited to E2500.00 following a serious motor vehicle accident;",
        "Windscreen cover is included where the Insured is only required to contribute nominal excess of 10% of the cost of repairs.",
    ]

    for point in bullet_points:
        bullet_paragraph = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph)

    Content_paragraph12 = doc.add_heading(
        "Motor Vehicle Warranty Reinstatement Extension", level=2
    )
    set_font_style(
        Content_paragraph12, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    Content_paragraph13 = doc.add_paragraph(
        "A new extension called Motor vehicle warranty reinstatement is included. Following accidental damage, the insured is covered for costs incurred on warranty reinstatement up to a limit of E5, 000.00."
    )

    Content_paragraph14 = doc.add_heading(
        "Optional Motor vehicle key cover extension", level=2
    )
    set_font_style(
        Content_paragraph14, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    Content_paragraph15 = doc.add_paragraph(
        "Following the accidental loss of keys, the cover is extended to include costs of replacement keys and other auxiliary costs incurred in re-programming the vehicle key. An excess of 10% of claim minimum E500.00 will apply."
    )

    Content_paragraph16 = doc.add_paragraph(
        "This extension shall be optional at an additional premium, and the options are as follows:\n"
    )
    set_font_style(Content_paragraph16, bold=True)

    # Add the table with Limit and Premium
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"

    # Add the header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Limit"
    hdr_cells[1].text = "Premium"

    # Make the header text bold
    for cell in hdr_cells:
        for run in cell.paragraphs[0].runs:
            run.font.name = "Cambria"
            run.font.size = Pt(11)
            run.bold = True  # Bold the header text

    # Add the data rows
    data = [["E10, 000.00", "E650.00"], ["E15, 000.00", "E950.00"]]

    for row_data in data:
        row = table.add_row().cells
        row[0].text = row_data[0]
        row[0].paragraphs[0].runs[0].font.name = "Cambria"
        row[0].paragraphs[0].runs[0].font.size = Pt(11)

        row[1].text = row_data[1]
        row[1].paragraphs[0].runs[0].font.name = "Cambria"
        row[1].paragraphs[0].runs[0].font.size = Pt(11)

    Content_paragraph17 = doc.add_heading("Optional Car Hire Extensions", level=2)
    set_font_style(
        Content_paragraph17, bold=True, underline=True, color=RGBColor(0, 0, 0)
    )

    Content_paragraph18 = doc.add_paragraph("Our car hire extension is divided into;")

    Content_paragraph19 = doc.add_paragraph()
    # Sub-header
    part1_run = Content_paragraph19.add_run("Standard Option\n")
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # standard option definition
    part2_run = Content_paragraph19.add_run(
        "This will enable policyholder whose vehicle has been involved in an accident, stolen or written off to a vehicle for a duration of 21 days while a claim is being processed.\n"
    )
    part2_run.font.name = "Cambria"
    part2_run.font.size = Pt(11)

    # Sub-header
    part3_run = Content_paragraph19.add_run("Enhanced Option\n")
    part3_run.font.name = "Cambria"
    part3_run.font.size = Pt(11)
    part3_run.bold = True

    # Enhanced option definition
    part4_run = Content_paragraph19.add_run(
        "In addition to what is provided under our standard option. We shall hire a vehicle for the policyholder following these events;"
    )
    part4_run.font.name = "Cambria"
    part4_run.font.size = Pt(11)

    # Add the bullet points
    bullet_points = [
        "For 5 days following mechanical or electrical failure of the insured vehicle.",
        "For 2 days in the event, the professional repairer or service garage is unable to finish routine services of the insured vehicle",
    ]

    for point in bullet_points:
        bullet_paragraph2 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph2)

    Content_paragraph20 = doc.add_paragraph(
        "Other terms and conditions of this optional extension are fully detailed in the policy wording."
    )

    Content_paragraph21 = doc.add_paragraph("Additional premiums are as follows;")
    set_font_style(Content_paragraph21, bold=True)

    # Add the table with Limit and Premium
    # Add table
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    headers = [
        "Vehicle description",
        "Standard option monthly premium",
        "Enhanced option monthly premium",
        "Days",
    ]

    hdr_cells = table.rows[0].cells

    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.font.name = "Cambria"
        run.font.size = Pt(11)
        run.bold = True

    # Add the data rows
    data = [
        ["Manual 1.4/1.6", "E105.00", "E125.00", "21"],
        ["Automatic 1.6", "E155.00", "E175.00", "21"],
    ]

    for row_data in data:
        row = table.add_row().cells
        for i, cell_value in enumerate(row_data):
            row[i].text = cell_value
            run = row[i].paragraphs[0].runs[0]
            run.font.name = "Cambria"
            run.font.size = Pt(11)

    Content_paragraph22 = doc.add_heading("SwaziMed Discount Extension", level=2)
    set_font_style(Content_paragraph22, bold=True, underline=True)

    Content_paragraph23 = doc.add_paragraph(
        "In the event you have your medical aid covered with Swazimed kindly provide us with a certificate confirming same to earn a discounted premium."
    )

    Content_paragraph24 = doc.add_paragraph()
    # Contact
    part1_run = Content_paragraph24.add_run(
        "Kindly provide us with current contact details e.g. Cell Number, E-mail Address etc.\n"
    )
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # trauma
    part2_run = Content_paragraph24.add_run(
        "Please advise our office on any amendments required regarding the basis of cover, we will gladly assist you."
    )
    part2_run.font.name = "Cambria"
    part2_run.font.size = Pt(11)

    Content_paragraph25 = doc.add_paragraph("Warm Regards\n")
    set_font_style(Content_paragraph25, bold=True)

    Content_paragraph26 = doc.add_paragraph()

    # Department
    part1_run = Content_paragraph26.add_run("UNDERWRITING TEAM\n")
    part1_run.font.name = "Cambria"
    part1_run.font.size = Pt(11)
    part1_run.bold = True

    # email
    part2_run = Content_paragraph26.add_run("e-mail: underwriting@sric.sz")
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
