#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script to find 'patient_enroules.xlsx' files in a nested directory structure,
rename them sequentially, and copy them to a new 'nouveau_patients' directory.
"""

import os
import shutil


def process_patient_files():
    """
    Main function to process the patient files.
    - Navigates a specific directory structure: ./output/<folder>/<subfolder>/
    - Finds 'patient_enroules.xlsx'.
    - Renames it to 'patient_enroules(i).xlsx' in its original location.
    - Copies the renamed file to './nouveau_patients/'.
    - Increments 'i' for each successful operation.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output")
    nouveau_patients_dir = os.path.join(script_dir, "nouveau_resultats")

    # --- 1. Create the destination folder if it doesn't exist ---
    try:
        os.makedirs(nouveau_patients_dir, exist_ok=True)
        print(f"Destination directory '{nouveau_patients_dir}' is ready.")
    except OSError as e:
        print(f"Erreur: Impossible de créer le dossier de destination '{nouveau_patients_dir}': {e}")
        return  # Stop if destination can't be created

    # --- 2. Check if the 'output' directory exists ---
    if not os.path.isdir(output_dir):
        print(f"Erreur: Le dossier 'output' n'a pas été trouvé dans '{script_dir}'.")
        print("Veuillez vous assurer que le dossier 'output' existe et contient les sous-dossiers attendus.")
        return

    print(f"Recherche de fichiers dans le dossier: '{output_dir}'")
    file_counter = 1  # This is 'i'

    # --- 3. Iterate through the directory structure ---
    # Level 1: Folders directly inside 'output'
    for first_level_folder_name in os.listdir(output_dir):
        first_level_folder_path = os.path.join(output_dir, first_level_folder_name)

        if os.path.isdir(first_level_folder_path):
            # print(f"  Analyse du dossier de premier niveau: '{first_level_folder_name}'") # Optional: for more verbose logging
            # Level 2: Subfolders inside each 'first_level_folder'
            for second_level_folder_name in os.listdir(first_level_folder_path):
                second_level_folder_path = os.path.join(first_level_folder_path, second_level_folder_name)

                if os.path.isdir(second_level_folder_path):
                    # print(f"    Analyse du sous-dossier: '{second_level_folder_name}'") # Optional
                    original_filename = "resultats.xlsx"
                    original_filepath = os.path.join(second_level_folder_path, original_filename)

                    # --- 4. Find 'patient_enroules.xlsx' ---
                    if os.path.isfile(original_filepath):
                        print(f"    Fichier trouvé: '{original_filepath}'")

                        # --- 5. Rename the file in its original location ---
                        new_filename_in_source = f"resultats({file_counter}).xlsx"
                        renamed_filepath_in_source = os.path.join(second_level_folder_path, new_filename_in_source)

                        try:
                            os.rename(original_filepath, renamed_filepath_in_source)
                            print(f"      Fichier renommé en: '{renamed_filepath_in_source}'")
                        except OSError as e:
                            print(
                                f"      Erreur: Impossible de renommer '{original_filepath}' en '{renamed_filepath_in_source}': {e}")
                            print(
                                f"      Il est possible qu'un fichier nommé '{new_filename_in_source}' existe déjà ou qu'il y ait un problème de permissions.")
                            continue  # Skip to the next file/folder if rename fails

                        # --- 6. Copy the renamed file to 'nouveau_patients' ---
                        destination_filepath = os.path.join(nouveau_patients_dir, new_filename_in_source)
                        try:
                            shutil.copy2(renamed_filepath_in_source, destination_filepath)  # copy2 preserves metadata
                            print(f"      Fichier copié vers: '{destination_filepath}'")
                            file_counter += 1  # Increment counter only on successful rename AND copy
                        except IOError as e:
                            print(
                                f"      Erreur: Impossible de copier '{renamed_filepath_in_source}' vers '{destination_filepath}': {e}")
                            # Optional: decide if you want to try and rename the source file back if copy fails
                            # For now, we'll leave the source file renamed even if copy fails.
                    # else: # Optional: if you want to log when the target file is not found in a subfolder
                    # print(f"    '{original_filename}' non trouvé dans '{second_level_folder_path}'.")

    if file_counter == 1:
        print(
            "\nAucun fichier 'recettes.xlsx' n'a été trouvé et traité selon la structure de dossiers attendue.")
    else:
        print(f"\nTraitement terminé. {file_counter - 1} fichier(s) ont été renommés et copiés.")


if __name__ == "__main__":
    process_patient_files()