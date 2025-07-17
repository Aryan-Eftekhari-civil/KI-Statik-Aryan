import pandas as pd
import os
import json
from pathlib import Path
from difflib import SequenceMatcher


def similarity(a, b):
    """Calculate similarity between two strings"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def clean_text_for_matching(text):
    """Clean text for better matching by removing special characters and normalizing"""
    import re
    # Convert to string and normalize
    text = str(text)

    # Replace underscores, hyphens, and dots with spaces for better matching
    text = text.replace('_', ' ').replace('-', ' ').replace('.', ' ')

    # Remove special characters but keep spaces, letters, and numbers
    cleaned = re.sub(r'[^\w\s]', ' ', text)

    # Replace multiple spaces with single space and normalize case
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().lower()

    return cleaned


def extract_meaningful_words(text, min_length=3):
    """Extract meaningful words from text, filtering out common short words"""
    import re

    # Common words to ignore (expand this list as needed)
    stop_words = {'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can', 'had', 'her', 'was', 'one', 'our',
                  'out', 'day', 'get', 'has', 'him', 'his', 'how', 'its', 'may', 'new', 'now', 'old', 'see', 'two',
                  'way', 'who', 'boy', 'did', 'man', 'end', 'few', 'got', 'let', 'put', 'say', 'she', 'too', 'use'}

    # Clean and split into words
    cleaned = clean_text_for_matching(text)
    words = re.findall(r'\w+', cleaned.lower())

    # Filter meaningful words
    meaningful_words = [word for word in words
                        if len(word) >= min_length
                        and word not in stop_words
                        and not word.isdigit()]  # Remove pure numbers

    return meaningful_words


def calculate_word_overlap_score(pdf_words, doc_words):
    """Calculate how many meaningful words overlap between PDF and document names"""
    if not pdf_words or not doc_words:
        return 0.0

    pdf_set = set(pdf_words)
    doc_set = set(doc_words)

    intersection = pdf_set.intersection(doc_set)
    union = pdf_set.union(doc_set)

    if not union:
        return 0.0

    # Jaccard similarity with bonus for longer matches
    jaccard_score = len(intersection) / len(union)

    # Bonus for having multiple word matches
    match_bonus = min(len(intersection) * 0.1, 0.3)

    return min(jaccard_score + match_bonus, 1.0)


def fuzzy_word_match(word1, word2, threshold=0.85):
    """Check if two words are similar enough (handles typos)"""
    if len(word1) < 4 or len(word2) < 4:
        return word1 == word2  # Exact match for short words

    return similarity(word1, word2) >= threshold


def rename_pdfs_with_document_numbers(excel_file_path, pdf_folder_path, sheet_name=0):
    """
    Match PDF filenames with document names in column H,
    then rename using corresponding document numbers from column V
    """

    try:
        # Read Excel file
        print("Reading Excel file...")
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        # Extract document names (column H) and numbers (column V) from the same rows
        print("Extracting document names and numbers...")

        # Create mapping from matching rows only
        name_to_number = {}
        valid_pairs = 0

        print("Document mappings found:")
        print("-" * 60)

        for index, row in df.iterrows():
            doc_name = row.iloc[7] if pd.notna(row.iloc[7]) else None  # Column H (index 7)
            doc_number = row.iloc[21] if pd.notna(row.iloc[21]) else None  # Column V (index 21)

            if doc_name is not None and doc_number is not None:
                # Clean the values
                clean_name = str(doc_name).strip()
                clean_number = str(doc_number).strip()

                # Handle numeric document numbers that might be read as floats
                if '.' in clean_number and clean_number.replace('.', '').replace('-', '').isdigit():
                    try:
                        clean_number = str(int(float(clean_number)))
                    except:
                        pass  # Keep original if conversion fails

                name_to_number[clean_name] = clean_number
                print(f"Row {index + 2}: '{clean_name}' -> Document #{clean_number}")
                valid_pairs += 1

        print(f"\nTotal valid document pairs: {valid_pairs}")

        if valid_pairs == 0:
            print("❌ No document name-number pairs found!")
            print("Please check that:")
            print("- Column H contains document names")
            print("- Column V contains document numbers")
            print("- Both columns have data in the same rows")
            return

        # Get PDF files
        pdf_folder = Path(pdf_folder_path)
        if not pdf_folder.exists():
            print(f"Error: PDF folder '{pdf_folder_path}' does not exist")
            return

        pdf_files = list(pdf_folder.rglob("*.pdf"))

        if not pdf_files:
            print("No PDF files found in the specified folder")
            return

        print(f"\nFound {len(pdf_files)} PDF files:")
        for pdf in pdf_files:
            print(f"  - {pdf.relative_to(pdf_folder)}")

        # Match PDF files with document names
        print(f"\nMatching PDF files with document names:")
        print("=" * 70)

        matches = []
        unmatched_pdfs = []

        for pdf_file in pdf_files:
            pdf_name = pdf_file.stem  # filename without .pdf extension
            matched = False
            best_match = None
            best_score = 0
            best_doc_number = None
            best_match_type = None
            best_match_details = ""

            print(f"\nAnalyzing PDF: '{pdf_file.relative_to(pdf_folder)}'")

            # Extract meaningful words from PDF name
            pdf_words = extract_meaningful_words(pdf_name)
            clean_pdf_name = clean_text_for_matching(pdf_name)

            # Try different matching strategies (ordered by precision)
            for doc_name, doc_number in name_to_number.items():
                clean_doc_name = clean_text_for_matching(doc_name)
                doc_words = extract_meaningful_words(doc_name)

                # Strategy 1: Exact match (case-insensitive, normalized)
                if clean_pdf_name == clean_doc_name:
                    print(f"  ✓ EXACT MATCH: '{doc_name}' -> Document #{doc_number}")
                    print(f"    PDF normalized: '{clean_pdf_name}'")
                    print(f"    Doc normalized: '{clean_doc_name}'")
                    matches.append((pdf_file, doc_number, doc_name, "exact", "Perfect normalized match"))
                    matched = True
                    break

                # Strategy 2: High word overlap (most meaningful matches)
                word_overlap_score = calculate_word_overlap_score(pdf_words, doc_words)
                if word_overlap_score > 0.7 and len(pdf_words) >= 2 and len(doc_words) >= 2:
                    common_words = set(pdf_words).intersection(set(doc_words))
                    details = f"Common words: {', '.join(sorted(common_words))}"
                    print(
                        f"  ✓ HIGH WORD OVERLAP: '{doc_name}' (score: {word_overlap_score:.3f}) -> Document #{doc_number}")
                    print(f"    {details}")
                    if word_overlap_score > best_score:
                        best_score = word_overlap_score
                        best_match = doc_name
                        best_doc_number = doc_number
                        best_match_type = "word_overlap"
                        best_match_details = details

                # Strategy 3: Substring match with high precision requirements
                elif len(clean_doc_name) > 8:  # Only for longer document names
                    if clean_doc_name in clean_pdf_name:
                        # Require substantial overlap (at least 60% of the shorter string)
                        overlap_ratio = len(clean_doc_name) / max(len(clean_pdf_name), len(clean_doc_name))
                        if overlap_ratio > 0.6:
                            score = overlap_ratio * 0.8  # Lower than word overlap
                            print(
                                f"  ~ SUBSTRING MATCH: PDF contains '{doc_name}' (score: {score:.3f}) -> Document #{doc_number}")
                            if score > best_score:
                                best_score = score
                                best_match = doc_name
                                best_doc_number = doc_number
                                best_match_type = "substring"
                                best_match_details = f"Overlap ratio: {overlap_ratio:.3f}"

                    elif clean_pdf_name in clean_doc_name and len(clean_pdf_name) > 8:
                        overlap_ratio = len(clean_pdf_name) / max(len(clean_pdf_name), len(clean_doc_name))
                        if overlap_ratio > 0.6:
                            score = overlap_ratio * 0.75  # Even lower score
                            print(
                                f"  ~ REVERSE MATCH: '{doc_name}' contains PDF name (score: {score:.3f}) -> Document #{doc_number}")
                            if score > best_score:
                                best_score = score
                                best_match = doc_name
                                best_doc_number = doc_number
                                best_match_type = "reverse"
                                best_match_details = f"Overlap ratio: {overlap_ratio:.3f}"

                # Strategy 4: Fuzzy word matching (handles typos in individual words)
                elif len(pdf_words) >= 2 and len(doc_words) >= 2:
                    fuzzy_matches = 0
                    total_possible = min(len(pdf_words), len(doc_words))

                    for pdf_word in pdf_words:
                        for doc_word in doc_words:
                            if fuzzy_word_match(pdf_word, doc_word):
                                fuzzy_matches += 1
                                break

                    fuzzy_score = fuzzy_matches / max(len(pdf_words), len(doc_words))
                    if fuzzy_score > 0.6 and fuzzy_matches >= 2:  # At least 2 words must match
                        score = fuzzy_score * 0.7  # Lower than exact matches
                        print(
                            f"  ~ FUZZY WORD MATCH: '{doc_name}' (score: {score:.3f}, {fuzzy_matches}/{total_possible} words) -> Document #{doc_number}")
                        if score > best_score:
                            best_score = score
                            best_match = doc_name
                            best_doc_number = doc_number
                            best_match_type = "fuzzy_words"
                            best_match_details = f"{fuzzy_matches} words matched"

                # Strategy 5: Overall similarity (only for very high similarity)
                else:
                    sim_score = similarity(clean_pdf_name, clean_doc_name)
                    if sim_score > 0.85 and len(clean_pdf_name) > 6 and len(clean_doc_name) > 6:  # Much stricter
                        score = sim_score * 0.6  # Lowest priority
                        print(
                            f"  ~ HIGH SIMILARITY: '{doc_name}' (similarity: {sim_score:.3f}) -> Document #{doc_number}")
                        if score > best_score:
                            best_score = score
                            best_match = doc_name
                            best_doc_number = doc_number
                            best_match_type = "high_similarity"
                            best_match_details = f"String similarity: {sim_score:.3f}"

            # Use best match only if it meets minimum quality threshold
            if not matched and best_match and best_score > 0.5:  # Higher threshold
                print(
                    f"  ✓ BEST MATCH: '{best_match}' ({best_match_type}, score: {best_score:.3f}) -> Document #{best_doc_number}")
                if best_match_details:
                    print(f"    {best_match_details}")
                matches.append((pdf_file, best_doc_number, best_match, best_match_type, best_match_details))
                matched = True

            if not matched:
                print(f"  ✗ NO MATCH FOUND (highest score: {best_score:.3f})")
                if best_match:
                    print(f"    Best candidate was: '{best_match}' but score too low")
                unmatched_pdfs.append(pdf_file)

        # Show summary
        print(f"\n" + "=" * 70)
        print("MATCHING SUMMARY:")
        print("=" * 70)

        if matches:
            print(f"\nFiles to be renamed ({len(matches)}):")
            for i, match_data in enumerate(matches, 1):
                if len(match_data) == 5:  # New format with details
                    pdf_file, doc_number, matched_name, match_type, details = match_data
                else:  # Old format compatibility
                    pdf_file, doc_number, matched_name, match_type = match_data
                    details = ""

                new_name = f"{doc_number}{pdf_file.stem}.pdf"
                print(f"  {i:2d}. '{pdf_file.relative_to(pdf_folder)}' -> '{new_name}'")
                print(f"      Matched with: '{matched_name}' ({match_type})")
                if details:
                    print(f"      Details: {details}")
                print()

        if unmatched_pdfs:
            print(f"\nFiles with no match ({len(unmatched_pdfs)}):")
            for pdf_file in unmatched_pdfs:
                print(f"  '{pdf_file.relative_to(pdf_folder)}'")

        if not matches:
            print("No files can be renamed - no matches found.")
            print("\nDebugging tips:")
            print("1. Run option 1 to see what document names are in column H")
            print("2. Run option 2 to see your PDF filenames")
            print("3. Check if PDF names are similar to document names in Excel")
            print("4. Document names and PDF names should have some common words/text")
            return

        # Ask for confirmation
        proceed = input(f"\nProceed with renaming {len(matches)} files? (y/N): ").strip().lower()
        if proceed not in ['y', 'yes']:
            print("Operation cancelled.")
            return

        # Perform renaming
        print(f"\nRenaming files...")
        renamed_count = 0
        rename_history = {}  # Track original names and paths for undo functionality

        for match_data in matches:
            if len(match_data) == 5:  # New format
                pdf_file, doc_number, matched_name, match_type, details = match_data
            else:  # Old format compatibility
                pdf_file, doc_number, matched_name, match_type = match_data

            new_filename = f"{doc_number}{pdf_file.stem}.pdf"
            new_path = pdf_file.parent / new_filename

            # Check if target file already exists
            if new_path.exists():
                print(f"  ⚠ Skipping '{pdf_file.relative_to(pdf_folder)}' - '{new_filename}' already exists")
                continue

            try:
                # Save original path relative to PDF folder root and new filename for undo functionality
                relative_original_path = str(pdf_file.relative_to(pdf_folder))
                relative_new_path = str(new_path.relative_to(pdf_folder))

                rename_history[relative_new_path] = relative_original_path

                pdf_file.rename(new_path)
                print(f"  ✓ '{pdf_file.relative_to(pdf_folder)}' -> '{new_filename}'")
                renamed_count += 1
            except Exception as e:
                print(f"  ✗ Error renaming '{pdf_file.relative_to(pdf_folder)}': {e}")

        # Save rename history to a backup file
        if rename_history:
            backup_file = Path(pdf_folder_path) / "rename_backup.json"
            try:
                # Load existing backup if it exists
                existing_backup = {}
                if backup_file.exists():
                    with open(backup_file, 'r', encoding='utf-8') as f:
                        existing_backup = json.load(f)

                # Merge with new renames
                existing_backup.update(rename_history)

                # Save updated backup
                with open(backup_file, 'w', encoding='utf-8') as f:
                    json.dump(existing_backup, f, ensure_ascii=False, indent=2)

                print(f"\n📁 Backup file saved: {backup_file}")
                print("   (This file contains original paths for undo functionality)")

            except Exception as e:
                print(f"⚠ Warning: Could not save backup file: {e}")

        print(f"\n🎉 Successfully renamed {renamed_count} out of {len(matches)} files!")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()


def preview_excel_data(excel_file_path, sheet_name=0):
    """
    Preview the Excel data to verify column H and V contents
    """
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        print("EXCEL FILE PREVIEW")
        print("=" * 50)
        print(f"Total rows: {len(df)}")
        print(f"Total columns: {len(df.columns)}")

        print(f"\nColumn H (Document Names) - showing first 15 entries:")
        col_h = df.iloc[:, 7]  # Column H (index 7)
        for i, value in enumerate(col_h.head(15)):
            if pd.notna(value):
                print(f"  Row {i + 2}: '{value}'")
            else:
                print(f"  Row {i + 2}: <empty>")

        print(f"\nColumn V (Document Numbers) - showing first 15 entries:")
        col_u = df.iloc[:, 21]  # Column V (index 21)
        for i, value in enumerate(col_u.head(15)):
            if pd.notna(value):
                print(f"  Row {i + 2}: '{value}' ({type(value).__name__})")
            else:
                print(f"  Row {i + 2}: <empty>")

        # Count non-empty pairs
        valid_pairs = 0
        for i in range(len(df)):
            if pd.notna(col_h.iloc[i]) and pd.notna(col_u.iloc[i]):
                valid_pairs += 1

        print(f"\nValid document name-number pairs: {valid_pairs}")

    except Exception as e:
        print(f"Error reading Excel file: {e}")


def list_pdf_files(pdf_folder_path):
    """
    List all PDF files in the folder recursively
    """
    try:
        pdf_folder = Path(pdf_folder_path)
        pdf_files = list(pdf_folder.rglob("*.pdf"))

        print("PDF FILES IN FOLDER (including subfolders)")
        print("=" * 45)

        if not pdf_files:
            print("No PDF files found.")
            return

        # Group files by folder for better organization
        files_by_folder = {}
        for pdf_file in pdf_files:
            folder_path = pdf_file.parent.relative_to(pdf_folder)
            if folder_path == Path('.'):
                folder_key = "Root folder"
            else:
                folder_key = str(folder_path)

            if folder_key not in files_by_folder:
                files_by_folder[folder_key] = []
            files_by_folder[folder_key].append(pdf_file.name)

        # Display files organized by folder
        for folder, files in sorted(files_by_folder.items()):
            print(f"\n📁 {folder}:")
            for i, filename in enumerate(sorted(files), 1):
                print(f"   {i:2d}. {filename}")

        print(f"\nTotal: {len(pdf_files)} PDF files across all folders")

    except Exception as e:
        print(f"Error listing PDF files: {e}")


def undo_renames(pdf_folder_path):
    """
    Undo PDF renames using the backup file that contains original paths
    """
    try:
        pdf_folder = Path(pdf_folder_path)
        if not pdf_folder.exists():
            print(f"Error: PDF folder '{pdf_folder_path}' does not exist")
            return

        # Look for the backup file
        backup_file = pdf_folder / "rename_backup.json"

        if not backup_file.exists():
            print("❌ No backup file found!")
            print(f"Expected backup file: {backup_file}")
            print("\nThe backup file is created automatically when you rename files.")
            print("Without this file, we cannot reliably undo the renames.")
            print("\nFor manual undo, you'll need to:")
            print("1. Identify the document number prefix in each filename")
            print("2. Remove it manually to restore the original name")
            return

        # Load the backup data
        try:
            with open(backup_file, 'r', encoding='utf-8') as f:
                backup_data = json.load(f)
        except Exception as e:
            print(f"❌ Error reading backup file: {e}")
            return

        if not backup_data:
            print("❌ Backup file is empty - no rename history found")
            return

        print("UNDO RENAMES - Using backup file")
        print("=" * 50)
        print(f"Backup file: {backup_file}")
        print(f"Found {len(backup_data)} rename records")

        # Check which files can actually be undone (search recursively)
        revertible_files = []
        missing_files = []

        for current_relative_path, original_relative_path in backup_data.items():
            current_path = pdf_folder / current_relative_path
            original_path = pdf_folder / original_relative_path

            if current_path.exists():
                revertible_files.append((current_path, original_path, current_relative_path, original_relative_path))
            else:
                missing_files.append((current_relative_path, original_relative_path))

        if missing_files:
            print(f"\n⚠ {len(missing_files)} files from backup not found (may have been moved/deleted):")
            for current_rel, original_rel in missing_files:
                print(f"  Missing: '{current_rel}' (was originally: '{original_rel}')")

        if not revertible_files:
            print("\n❌ No files can be reverted - all files from backup are missing")
            return

        print(f"\n✅ {len(revertible_files)} files can be reverted:")
        print("-" * 70)

        for i, (current_path, original_path, current_rel, original_rel) in enumerate(revertible_files, 1):
            print(f"{i:2d}. Current:  '{current_rel}'")
            print(f"    Original: '{original_rel}'")
            print()

        # Check for conflicts
        conflicts = []
        for current_path, original_path, current_rel, original_rel in revertible_files:
            if original_path.exists() and original_path != current_path:
                conflicts.append((current_rel, original_rel))

        if conflicts:
            print(f"⚠ WARNING: {len(conflicts)} potential conflicts detected:")
            for current_rel, original_rel in conflicts:
                print(f"  '{current_rel}' -> '{original_rel}' (target already exists)")
            print()

        # Ask for confirmation
        proceed = input(f"Proceed with reverting {len(revertible_files)} files? (y/N): ").strip().lower()
        if proceed not in ['y', 'yes']:
            print("Undo operation cancelled.")
            return

        # Perform the undo operation
        print(f"\nReverting files to original names and locations...")
        reverted_count = 0
        skipped_count = 0
        reverted_items = []

        for current_path, original_path, current_rel, original_rel in revertible_files:
            # Create target directory if it doesn't exist
            original_path.parent.mkdir(parents=True, exist_ok=True)

            # Skip if target already exists
            if original_path.exists() and original_path != current_path:
                print(f"  ⚠ Skipping '{current_rel}' - '{original_rel}' already exists")
                skipped_count += 1
                continue

            try:
                current_path.rename(original_path)
                print(f"  ✓ '{current_rel}' -> '{original_rel}'")
                reverted_count += 1
                reverted_items.append(current_rel)
            except Exception as e:
                print(f"  ✗ Error reverting '{current_rel}': {e}")

        # Update backup file to remove successfully reverted items
        if reverted_items:
            try:
                updated_backup = {k: v for k, v in backup_data.items() if k not in reverted_items}

                if updated_backup:
                    # Save updated backup
                    with open(backup_file, 'w', encoding='utf-8') as f:
                        json.dump(updated_backup, f, ensure_ascii=False, indent=2)
                    print(f"\n📁 Backup file updated ({len(updated_backup)} entries remaining)")
                else:
                    # Remove empty backup file
                    backup_file.unlink()
                    print(f"\n📁 Backup file removed (all items reverted)")

            except Exception as e:
                print(f"⚠ Warning: Could not update backup file: {e}")

        print(f"\n🎉 Successfully reverted {reverted_count} files!")
        if skipped_count > 0:
            print(f"⚠ Skipped {skipped_count} files due to conflicts")

    except Exception as e:
        print(f"Error during undo operation: {e}")
        import traceback
        traceback.print_exc()


# Main execution
if __name__ == "__main__":
    # File paths
    excel_file = r"C:\Users\Eftekhari-FCP\Desktop\EXCEL_TEST\Anlagenverzeichnis_BS4 Gleisanhebung_TEST.xlsx"
    pdf_folder = r"C:\Users\Eftekhari-FCP\Desktop\PDF_TESTS"

    print("PDF RENAMER - Match Column H names with PDF files")
    print("=" * 55)

    while True:
        print("\nChoose an option:")
        print("1. Preview Excel data (Column H & V)")
        print("2. List PDF files (including subfolders)")
        print("3. Rename PDF files")
        print("4. Undo renames (revert to original names)")
        print("5. Exit")

        choice = input("\nEnter your choice (1-5): ").strip()

        if choice == "1":
            preview_excel_data(excel_file)
        elif choice == "2":
            list_pdf_files(pdf_folder)
        elif choice == "3":
            rename_pdfs_with_document_numbers(excel_file, pdf_folder)
        elif choice == "4":
            undo_renames(pdf_folder)
        elif choice == "5":
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Please enter 1, 2, 3, 4, or 5.")