import pandas as pd
import os
from datetime import datetime

base_dir = os.path.dirname(os.path.abspath(__file__))
source_folder = os.path.join(base_dir, "golda")
correspondence_file = os.path.join(base_dir, "Tableau_Correspondance_Golda.xlsx")
marque_correspondence_file = os.path.join(
    base_dir, "Tableau_Correspondance_Marque.xlsx"
)


def remove_last_character(s):
    return s[:-1]


if not os.path.exists(source_folder):
    print(f"Source folder not found: {source_folder}")
else:
    if not os.path.exists(correspondence_file):
        print(f"Correspondence file not found: {correspondence_file}")
    elif not os.path.exists(marque_correspondence_file):
        print(f"Marque correspondence file not found: {marque_correspondence_file}")
    else:
        correspondence_df = pd.read_excel(correspondence_file)
        marque_correspondence_df = pd.read_excel(marque_correspondence_file)

        print("Columns in correspondence file:", correspondence_df.columns)
        print(
            "Columns in marque correspondence file:", marque_correspondence_df.columns
        )

        if (
            "Code Sous-famille GOLDA" not in correspondence_df.columns
            or "Code Sous-famille PRIMA" not in correspondence_df.columns
        ):
            print(
                "Required columns 'Code Sous-famille GOLDA' and 'Code Sous-famille PRIMA' not found in correspondence file."
            )
        elif (
            "Code Marque Golda" not in marque_correspondence_df.columns
            or "Code Marque Prima" not in marque_correspondence_df.columns
            or "Prefixe Tarif Golda" not in marque_correspondence_df.columns
        ):
            print(
                "Required columns 'Code Marque Golda', 'Code Marque Prima', and 'Prefixe Tarif Golda' not found in marque correspondence file."
            )
        else:
            correspondence_subfam_dict = dict(
                zip(
                    correspondence_df["Code Sous-famille GOLDA"],
                    correspondence_df["Code Sous-famille PRIMA"],
                )
            )
            marque_correspondence_dict = dict(
                zip(
                    marque_correspondence_df["Code Marque Golda"],
                    marque_correspondence_df["Code Marque Prima"],
                )
            )
            prefixe_tarif_dict = dict(
                zip(
                    marque_correspondence_df["Prefixe Tarif Golda"],
                    marque_correspondence_df["Code Marque Prima"],
                )
            )

            for root, dirs, files in os.walk(source_folder):
                for file in files:
                    if file.endswith(".xlsx"):
                        file_path = os.path.join(root, file)
                        print(f"Processing file: {file_path}")

                        df = pd.read_excel(file_path)

                        if len(df) > 2:
                            df = df.drop(index=[1, 2])

                        def is_buggy(row):
                            return pd.isnull(row["Prix_euro"]) or row["Prix_euro"] == ""

                        df = df[~df.apply(is_buggy, axis=1)]

                        if "Code_marque" in df.columns:
                            df["CODE MARQUE"] = df["Code_marque"].map(
                                marque_correspondence_dict
                            )
                            df["CODE MARQUE"] = df.apply(
                                lambda row: (
                                    prefixe_tarif_dict.get(
                                        row["Prefixe_tarif"], row["CODE MARQUE"]
                                    )
                                    if pd.isnull(row["CODE MARQUE"])
                                    else row["CODE MARQUE"]
                                ),
                                axis=1,
                            )

                        if "Poids" in df.columns:
                            df["Poids"] = df["Poids"] / 1000
                        if "Ref_fournisseur" in df.columns:
                            df["REF APPEL 1"] = df["Ref_fournisseur"]
                            df["REF APPEL 2"] = df["Ref_fournisseur"].str.replace(
                                r"[^a-zA-Z0-9]", "", regex=True
                            )

                        df["Code_sousfamille_NU"] = df["Code_sousfamille_NU"].map(
                            correspondence_subfam_dict
                        )
                        df["Code_famille_NU"] = (
                            df["Code_sousfamille_NU"].astype(str).str.slice(0, 3)
                        )

                        # Initialize the 'Consigne' column for original articles
                        df["Consigne"] = ""

                        # Create consignment articles
                        consignment_rows = []
                        consignment_refs = {}
                        for index, row in df.iterrows():
                            if (
                                pd.notnull(row["Montant_consigne"])
                                and row["Montant_consigne"] != ""
                            ):
                                montant_consigne = row["Montant_consigne"]
                                if montant_consigne not in consignment_refs:
                                    new_row = row.copy()
                                    if (
                                        pd.notnull(row["Reference_consigne"])
                                        and row["Reference_consigne"] != ""
                                    ):
                                        new_row["Ref_fournisseur"] = row[
                                            "Reference_consigne"
                                        ]
                                    else:
                                        montant_consigne_str = str(
                                            montant_consigne
                                        ).replace(".", "")
                                        if montant_consigne_str.endswith("0"):
                                            montant_consigne_str = (
                                                remove_last_character(
                                                    montant_consigne_str
                                                )
                                            )
                                        new_row["Ref_fournisseur"] = (
                                            row["Ref_fournisseur"]
                                            + "-"
                                            + montant_consigne_str
                                        )
                                    new_row["Prix_euro"] = montant_consigne
                                    new_row["Code_famille_NU"] = "CON"
                                    new_row["Code_sousfamille_NU"] = "CON99"
                                    new_row["Description"] = (
                                        "CONSIGNE " + row["Description"]
                                    )
                                    new_row["Description_courte"] = ""
                                    new_row["Code_marque"] = row["Code_marque"]
                                    new_row["Code_EAN"] = ""
                                    new_row["REF APPEL 1"] = ""
                                    new_row["REF APPEL 2"] = ""
                                    new_row["Poids"] = ""
                                    new_row["Code_remise"] = ""
                                    new_row["Consigne"] = (
                                        "O"  # Indicate that this is a consignment
                                    )
                                    consignment_rows.append(new_row)
                                    consignment_refs[montant_consigne] = new_row[
                                        "Ref_fournisseur"
                                    ]

                        # Append consignment rows to the dataframe
                        if consignment_rows:
                            df = pd.concat(
                                [df, pd.DataFrame(consignment_rows)], ignore_index=True
                            )

                        # Add consignment reference to original articles
                        df["Consigne_Ref"] = df.apply(
                            lambda row: (
                                consignment_refs.get(row["Montant_consigne"], "")
                                if pd.notnull(row["Montant_consigne"])
                                and row["Consigne"] != "O"
                                else ""
                            ),
                            axis=1,
                        )

                        columns_to_extract = [
                            "CODE MARQUE",
                            "Ref_fournisseur",
                            "Code_EAN",
                            "REF APPEL 1",
                            "REF APPEL 2",
                            "Description",
                            "Description_courte",
                            "Code_famille_NU",
                            "Code_sousfamille_NU",
                            "Poids",
                            "Prix_euro",
                            "Code_remise",
                            "Consigne",
                            "Consigne_Ref",
                        ]

                        extracted_columns = df[columns_to_extract]

                        new_column_names = {
                            "CODE MARQUE": "CODE MARQUE",
                            "Ref_fournisseur": "CODE ARTICLE",
                            "Code_EAN": "GENCODE",
                            "REF APPEL 1": "REF APPEL 1",
                            "REF APPEL 2": "REF APPEL 2",
                            "Description": "LIBELLE ARTICLE",
                            "Description_courte": "LIBELLE COMPLEMENTAIRE",
                            "Code_famille_NU": "CODE FAMILLE PRIMA",
                            "Code_sousfamille_NU": "CODE SOUS FAMILLE PRIMA",
                            "Poids": "POIDS",
                            "Prix_euro": "PRIX DE BASE",
                            "Code_remise": "CODE REMISE",
                            "Consigne": "CONSIGNE",
                            "Consigne_Ref": "CONSIGNE REF",
                        }
                        extracted_columns.rename(columns=new_column_names, inplace=True)

                        original_file_name = os.path.basename(file_path)
                        new_file_prefix = original_file_name[:3].upper()
                        current_date = datetime.now().strftime("%d-%m-%Y")

                        new_file_name = (
                            f"{new_file_prefix} TARIF AU {current_date}.xlsx"
                        )
                        output_file_path = os.path.join(
                            os.path.dirname(file_path), new_file_name
                        )

                        extracted_columns.to_excel(output_file_path, index=False)

                        print(f"Extracted columns saved to: {output_file_path}")
