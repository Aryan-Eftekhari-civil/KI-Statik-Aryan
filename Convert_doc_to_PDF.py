import os
import win32com.client

def convert_docm_to_pdf(folder_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".docx"):
            docm_path = os.path.join(folder_path, filename)

            # Change only the file extension to .pdf
            pdf_filename = os.path.splitext(filename)[0] + ".pdf"
            pdf_path = os.path.join(folder_path, pdf_filename)

            print(f"Converting: {filename} → {pdf_filename}")

            doc = word.Documents.Open(docm_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF
            doc.Close()

    word.Quit()
    print("All files converted.")


# Example usage:
folder = r"R:\A_Auftraege\2023\1156\B_Berichte-SV\20_Berichte\06-CSM_Dokumente_BS4\5_Erklärung des Vorschlagenden"
convert_docm_to_pdf(folder)


