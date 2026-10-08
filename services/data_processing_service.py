import pandas as pd
import os


def process_data(df: pd.DataFrame, output_folder: str):
    """
    Full renewal processing logic.
    Cleans, transforms, filters and saves Excel outputs.
    Returns dictionary of filtered DataFrames for letter generation.
    """

    df = df.copy()

    # -------------------------------------------------------
    # REORDER total_sum_insured COLUMN
    # -------------------------------------------------------
    if "total_sum_insured" in df.columns:
        cols = df.columns.tolist()
        cols.remove("total_sum_insured")
        cols.insert(-1, "total_sum_insured")
        df = df[cols]

    # -------------------------------------------------------
    # SPLIT RISK DESCRIPTION
    # -------------------------------------------------------
    def split_at_second_space(text):
        parts = str(text).split(" ", 2)
        if len(parts) < 3:
            parts += [""] * (3 - len(parts))
        return [" ".join(parts[:2]), parts[2]]

    if "risk_description" in df.columns:
        df_split = df["risk_description"].apply(split_at_second_space).apply(pd.Series)
        df_split.columns = ["RiskType", "risk_description"]
        df = pd.concat([df, df_split], axis=1)

    # -------------------------------------------------------
    # RENAME COLUMNS
    # -------------------------------------------------------
    df = df.rename(
        columns={
            "Insurance_ref": "Policy_Number",
            "total_sum_insured": "SumInsured",
        }
    )

    # -------------------------------------------------------
    # CLEAN STRINGS
    # -------------------------------------------------------
    df["RiskType"] = (
        df["RiskType"]
        .str.replace(",", " ", regex=False)
        .str.replace("-", " ", regex=False)
        .str.strip()
        .str.replace(" ", "_", regex=False)
    )

    df["Policy_Number"] = df["Policy_Number"].astype(str).str.strip()

    # -------------------------------------------------------
    # REMOVE STAMP DUTY
    # -------------------------------------------------------
    df = df[~df["RiskType"].str.lower().isin(["stamp_duty", "stampduty"])]

    # -------------------------------------------------------
    # GROUPING LOGIC
    # -------------------------------------------------------
    risk_types = df.groupby("Policy_Number")["RiskType"].agg(list).reset_index()
    risk_counts = (
        df.groupby("Policy_Number")["RiskType"].size().reset_index(name="RiskTypeCount")
    )

    df = df.merge(risk_types, on="Policy_Number", suffixes=("", "_all_risks"))
    df = df.merge(risk_counts, on="Policy_Number")

    df["AllRiskTypes"] = df["RiskType_all_risks"]
    df = df.drop(columns=["RiskType_all_risks"])

    unique_risk_types = (
        df.groupby("Policy_Number")["RiskType"]
        .apply(lambda x: list(set(x)))
        .reset_index(name="UniqueRiskTypes")
    )

    df = df.merge(unique_risk_types, on="Policy_Number")

    # -------------------------------------------------------
    # COMPUTE HIGHEST PRIVATE MOTOR SUM
    # -------------------------------------------------------
    max_private_motor = (
        df[df["RiskType"] == "Private_Motor"]
        .groupby("Policy_Number")["SumInsured"]
        .max()
    )

    df["HighestPrivateMotorSum"] = df["Policy_Number"].map(max_private_motor)

    # -------------------------------------------------------
    # DROP UNUSED COLUMNS
    # -------------------------------------------------------
    columns_to_drop = [
        "insurance_file_cnt",
        "insurance_holder_cnt",
        "product_id",
        "lead_agent_cnt",
        "Policy_stop_reason",
        "Client_stop_reason",
        "insurance_folder_cnt",
        "client_code",
        "Agent_stop_reason",
        "is_deleted",
        "is_in_transfer_mode",
        "is_true_monthly_policy",
        "anniversary_copy",
        "renewal_day_number",
        "anniversary_renewal_weeks",
        "put_on_next_instalment_renewal",
        "lead_allow_consolidated_commission",
        "sub_allow_consolidated_commission",
        "renewal_product_id",
        "original_product_id",
        "tmpautrenfac",
    ]

    existing_columns = [col for col in columns_to_drop if col in df.columns]
    df = df.drop(columns=existing_columns, errors="ignore")

    # -------------------------------------------------------
    # SAVE MASTER FILE
    # -------------------------------------------------------
    all_policies_path = os.path.join(output_folder, "All_policies.xlsx")
    df.to_excel(all_policies_path, index=False)

    # -------------------------------------------------------
    # FILTER DIRECT BUSINESS
    # -------------------------------------------------------
    df_direct = df[df["agent_name"].isnull()]
    df_direct.to_excel(os.path.join(output_folder, "Direct_Business.xlsx"), index=False)
    
    # -------------------------------------------------------
    # FILTER DIRECT BUSINESS AND SELECTED AGENTS
    # -------------------------------------------------------
    agents = [
"Khanyisile Patience Myeni", "Khanyisile Queeneth Dlamini",
"Khulekani Mzileni","Mlondie Dlamini",
"Ncamsile Nonchazelo Dlamini","Nhlanhla Lucky Ginindza",
"Nontokozo Zikalala","Sicelo Masuku","Zinhle Nolwazi Sithole"
"Zodwa Lukhele","Amanda Qondile Dlamini","Bheki Gwebu",
"Celumusa Methula","Fisokuhle Mnisi","Jabulani Nxumalo",
"Jabulile Glenrose Dlamini","Kenneth Nkambule","Lungile Precious Tsabedze",
"Mxolisi Erasmus Mkhatshwa","Mlondie Dlamini","Mlondi Vusi Mamba",
"Mumcy Mildred Ndzimandze","Musha Moses Nzima","Mxolisi Mdluli",
"Nolwazi Knowledge Dlamini","Nompumelelo Mavimbela","Nosizo Dlamini",
"Ntombifuthi Nozizwe Shabangu","Penuel  Malinga","Petros Methula",
"Philangaye Simelane","Sakhile Fortunate Gama","Sanele Patrick Nkambule",
"Saneliso Dlamini","Sibonelo Sibandze","Siboniso Dlamini",
"Sibusiso Lukhele","Sihle Dlamini","Sihle Pearl Mnisi","Siphiwe Fakudze"
"Sydney Nyembe","TA Maziya","Temahlungwane Nolwazi Mkhonta",
"Temalangeni Dlamini","Thembinkosi Phinda Phiri","Treasure Dlamini",
"Webster Simelane","William Mavuso","Wilson Simelane",
"Winile Aah Dlamini","Zwelithini Gamedze"]

    df_direct_with_agents = df[df["agent_name"].isnull() |df["agent_name"].isin(agents)]
    df_direct_with_agents.to_excel(os.path.join(output_folder, "Direct_With_Agents_Business.xlsx"), index=False)


    # -------------------------------------------------------
    # REMOVE DUPLICATES HELPER
    # -------------------------------------------------------
    def remove_duplicates(df):
        return df.drop_duplicates(subset="Policy_Number")

    # -------------------------------------------------------
    # FILTERING LOGIC
    # -------------------------------------------------------
    DPA_list_1 = ["Private_Motor", "Domestic_Property"]
    DPA_list_2 = ["Private_Motor", "Domestic_Property", "Domestic_Personal"]

    pm_less_300k = remove_duplicates(
        df_direct_with_agents[
            (df_direct_with_agents["product_description"] == "Domestic Protector")
            & (
                df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(["Private_Motor"])
                )
            )
            & (df_direct_with_agents["HighestPrivateMotorSum"] < 300000)
        ]
    )

    pm_above_300k = remove_duplicates(
        df_direct_with_agents[
            (df_direct_with_agents["product_description"] == "Domestic Protector")
            & (
                df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(["Private_Motor"])
                )
            )
            & (df_direct_with_agents["HighestPrivateMotorSum"] >= 300000)
        ]
    )

    dpa_above_300k = remove_duplicates(
        df_direct_with_agents[
            (df_direct_with_agents["product_description"] == "Domestic Protector")
            & (
                df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(DPA_list_1)
                )
                | df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(DPA_list_2)
                )
            )
            & (df_direct_with_agents["HighestPrivateMotorSum"] >= 300000)
        ]
    )

    dpa_less_300k = remove_duplicates(
        df_direct_with_agents[
            (df_direct_with_agents["product_description"] == "Domestic Protector")
            & (
                df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(DPA_list_1)
                )
                | df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(DPA_list_2)
                )
            )
            & (df_direct_with_agents["HighestPrivateMotorSum"] < 300000)
        ]
    )

    houseowners_only = remove_duplicates(
        df_direct_with_agents[
            (df_direct_with_agents["product_description"] == "Domestic Protector")
            & (
                df_direct_with_agents["UniqueRiskTypes"].apply(
                    lambda x: sorted(x) == sorted(["Domestic_Property"])
                )
            )
        ]
    )

    multimark = remove_duplicates(
        df_direct_with_agents[df_direct_with_agents["product_description"] == "MultiMark"]
    )
    special_fleet = remove_duplicates(
        df_direct_with_agents[df_direct_with_agents["product_description"] == "Special Fleet"]
    )
    workmens = remove_duplicates(
        df_direct_with_agents[df_direct_with_agents["product_description"] == "Workmen's Compensation"]
    )

    # -------------------------------------------------------
    # SAVE TO SERVICE FOLDERS
    # -------------------------------------------------------
    def save_to_folder(df, folder_name):
        if not df.empty:
            folder_path = os.path.join(output_folder, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            df.to_excel(os.path.join(folder_path, f"{folder_name}.xlsx"), index=False)

    save_to_folder(pm_less_300k, "PM less 300k")
    save_to_folder(pm_above_300k, "PM above 300K")
    save_to_folder(dpa_above_300k, "DPA above 300k")
    save_to_folder(dpa_less_300k, "DPA less 300k")
    save_to_folder(houseowners_only, "HouseOwners only")
    save_to_folder(multimark, "MultiMark")
    save_to_folder(special_fleet, "Special Fleet")
    save_to_folder(workmens, "Workmens Compensation")

    # -------------------------------------------------------
    # RETURN DICTIONARY FOR LETTER GENERATION
    # -------------------------------------------------------
    return {
        "PM less 300k": pm_less_300k,
        "PM above 300K": pm_above_300k,
        "DPA above 300k": dpa_above_300k,
        "DPA less 300k": dpa_less_300k,
        "HouseOwners only": houseowners_only,
        "MultiMark": multimark,
        "Special Fleet": special_fleet,
        "Workmens Compensation": workmens,
    }
