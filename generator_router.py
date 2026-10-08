# =====================================================
# generator_router.py  (SAFE VERSION)
# =====================================================

from letter_generators.houseowners import generate as houseowners_generate
from letter_generators.pm_less_300k import generate as pm_less_generate
from letter_generators.pm_above_300k import generate as pm_above_generate
from letter_generators.dpa_less_300k import generate as dpa_less_generate
from letter_generators.dpa_above_300k import generate as dpa_above_generate
from letter_generators.special_fleet import generate as special_fleet_generate
from letter_generators.multimark import generate as multimark_generate
from letter_generators.workmens_compensation import generate as workmens_generate

from services.word_to_pdf_service import convert_docx_to_pdf


def route_letter_generator(
    service,
    row,
    output_folder,
    header_image_path,
    footer_image_path,
):
    service_map = {
        "PM less 300k": pm_less_generate,
        "PM above 300K": pm_above_generate,
        "DPA less 300k": dpa_less_generate,
        "DPA above 300k": dpa_above_generate,
        "HouseOwners only": houseowners_generate,
        "MultiMark": multimark_generate,
        "Special Fleet": special_fleet_generate,
        "Workmens Compensation": workmens_generate,
    }

    generator = service_map.get(service)

    if not generator:
        raise ValueError(f"No generator implemented for {service}")

    # -------------------------------------------------
    # EXTRACT VALUES HERE (ONE PLACE ONLY)
    # -------------------------------------------------
    client_name = str(row.get("client_name", ""))
    address1 = str(row.get("address1", ""))
    address2 = str(row.get("address2", ""))
    policy_number = str(row.get("Policy_Number", "UNKNOWN"))
    cover_start_date = str(row.get("cover_start_date", ""))

    # -------------------------------------------------
    # CALL GENERATOR (NO ROW ANYMORE)
    # -------------------------------------------------
    docx_path = generator(
        client_name=client_name,
        address1=address1,
        address2=address2,
        Policy_Number=policy_number,
        cover_start_date=cover_start_date,
        output_folder=output_folder,
        header_image_path=header_image_path,
        footer_image_path=footer_image_path,
    )

    try:
        convert_docx_to_pdf(docx_path)
    except Exception as e:
        print(f"PDF conversion failed: {e}")

    return docx_path
