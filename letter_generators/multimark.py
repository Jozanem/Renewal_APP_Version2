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
        f"POLICY RENEWAL: {Policy_Number} – MULTIMARK III POLICY"
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

    Content_paragraph3 = doc.add_heading("Stated Benefits", level=2)
    set_font_style(Content_paragraph3, bold=True, color=RGBColor(0, 0, 0))

    # Add the bullet points
    bullet_points = [
        "It provides cover against any accident through external and visible means.",
        "Cover applies 24 hours a day and anywhere in the world.",
        "It was designed to complement any other accident insurance or life insurance the person may have.",
        "Employers generally take this cover to enhance staff benefits.",
        "The earnings of the top 3 earners in each organization must be declared separately.",
    ]

    for point in bullet_points:
        bullet_paragraph1 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph1)

    Content_paragraph4 = doc.add_heading("Group Personal Accident", level=2)
    set_font_style(Content_paragraph4, bold=True)

    # Add the bullet points
    bullet_points = [
        "Bodily injury caused by accidental, violent, external and visible means to any principal, partner, or employee of the insured.",
        "The Corporation will pay to the insured, on behalf of such person or his estate, the compensation in the event of accidental bodily injury to any such person directly and independently of all other causes resulting in death or disability.",
    ]
    for point in bullet_points:
        bullet_paragraph2 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph2)

    Content_paragraph5 = doc.add_heading("Business All Risk Section", level=2)
    set_font_style(Content_paragraph5, bold=True)

    bullet_points = [
        "Cover is against loss or damage to the whole or part of the item insured.",
        "Cover applies anywhere in the world, by any accident or misfortune not otherwise excluded",
    ]

    for point in bullet_points:
        bullet_paragraph3 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph3)

    Content_paragraph6 = doc.add_paragraph("There is no cover against the following:")

    bullet_points = [
        "Wear and tear or gradual deterioration.",
        "Mechanical, electronic or electrical breakdown or derangement.",
        "Theft from any unattended vehicle in custody of the insured or employee unless the vehicle is completely closed and securely locked, or the vehicle is in a securely locked building.",
        "Entry to or exit from such locked vehicle and building should be accompanied by forcible and violent entry or exit.",
    ]

    for point in bullet_points:
        bullet_paragraph4 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph4)

    Content_paragraph7 = doc.add_heading("Public Liability Section", level=2)
    set_font_style(Content_paragraph7, bold=True)

    bullet_points = [
        "The cover is against all third-party liabilities.",
        "These could be death, bodily injury or illness of persons.",
        "Includes physical loss or damage to tangible property.",
        "The cover is in the aggregate.",
        "Cover includes legal costs.",
    ]

    for point in bullet_points:
        bullet_paragraph5 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph5)

    Content_paragraph8 = doc.add_heading("Fidelity Guarantee", level=2)
    set_font_style(Content_paragraph8, bold=True)

    bullet_points = [
        "Loss of money and/or other property belonging to the insured or for which they are responsible, stolen by an insured employee.",
        "Direct financial loss sustained by the insured as a result of fraud or dishonesty of an insured employee which results in dishonest personal financial gain for the employee concerned.",
    ]

    for point in bullet_points:
        bullet_paragraph6 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph6)

    Content_paragraph9 = doc.add_heading("Money")
    set_font_style(Content_paragraph9, bold=True)

    Content_paragraph10 = doc.add_paragraph(
        "Loss of or damage to money. Cover is extended to include:"
    )

    bullet_points = ["Receptacles and clothing.", "Locks and keys ", "Skeletons keys"]

    for point in bullet_points:
        bullet_paragraph7 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph7)

    Content_paragraph11 = doc.add_heading("Electronic Equipment Section", level=2)
    set_font_style(Content_paragraph11, bold=True)

    Content_paragraph12 = doc.add_paragraph(
        "Cover is against physical loss of or damage to the property insured described in the schedule from any cause not hereinafter excluded whilst:"
    )

    bullet_points = [
        "At work or at rest anywhere within that part of the building occupied by the insured at the insured premises as described in the schedule.",
        "In transit including loading and unloading or whilst temporarily stored at any premises en route.",
        "Temporarily removed from that part of the building occupied by the insured at the premises described in the schedule to any other building.",
    ]

    for point in bullet_points:
        bullet_paragraph8 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph8)

    Content_paragraph13 = doc.add_paragraph(
        "There is no cover against but not limited to the following:"
    )

    bullet_points = [
        "Derangement unless accompanied by physical damage otherwise covered.",
        "Faults or defects known to the insured or employees at the time this insurance was arranged and not disclosed to the Corporation.",
        "Theft of property from vehicles unless, if left there overnight, the vehicle is in a securely locked building and there is forcible and violent entry and exit to the building.",
    ]

    for point in bullet_points:
        bullet_paragraph9 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph9)

    Content_paragraph14 = doc.add_heading("Buildings – Combined Section", level=2)
    set_font_style(Content_paragraph14, bold=True)

    Content_paragraph15 = doc.add_paragraph("Cover is against the following perils:")

    bullet_points = [
        "Fire, lightning & explosion.",
        "Storm, wind, water, hail or snow.",
        "Earthquake.",
        "Theft or attempted theft.",
        "Riot, strike, civil commotion.",
    ]

    for point in bullet_points:
        bullet_paragraph10 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph10)

    Content_paragraph16 = doc.add_paragraph(
        "Extensions include but not limited to the following:"
    )

    bullet_points = [
        "Policy has an in-built public liability extension limited to E1 000 000 in the aggregate.",
        "Capital additions-which covers alterations, additions and improvements not exceeding 15% of the sum insured but you must advise the Corporation of such and pay the additional premium.",
    ]

    for point in bullet_points:
        bullet_paragraph11 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph11)

    Content_paragraph17 = doc.add_heading("Office contents Section", level=2)
    set_font_style(Content_paragraph17, bold=True)

    bullet_points = [
        "Cover is for office contents of every description.",
        "Excludes laptops and cellphones.",
        "These could include office furniture, fixtures and fittings, computers and faxes.",
    ]

    for point in bullet_points:
        bullet_paragraph12 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph12)

    Content_paragraph17 = doc.add_paragraph("Cover is against the following:")

    bullet_points = [
        "Fire, lightning, explosion",
        "Earthquake",
        "Storm, wind, water, hail or snow",
        "Aircraft or other aerial devices",
        "Theft or attempted theft",
        "Riot, strike, civil commotion",
    ]

    for point in bullet_points:
        bullet_paragraph13 = doc.add_paragraph(point, style="ListBullet")
        set_font_style(bullet_paragraph13)

    Content_paragraph18 = doc.add_paragraph(
        "Kindly ensure that the sum insured represents the current replacement value of the property in order to avoid the application of average.\n"
    )
    set_font_style(Content_paragraph18, bold=True)

    Content_paragraph19 = doc.add_paragraph(
        "Please advise our office on any amendments required regarding the basis of cover, we will gladly assist you."
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
