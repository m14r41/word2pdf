import os
import win32com.client
from colorama import init, Fore

# Initialize colorama
init(autoreset=True)

# Print the ASCII Art for 'word2pdf'
print(Fore.RED + r"""

                       _ ____            _  __ 
__      _____  _ __ __| |___ \ _ __   __| |/ _|
\ \ /\ / / _ \| '__/ _` | __) | '_ \ / _` | |_
 \ V  V / (_) | | | (_| |/ __/| |_) | (_| |  _|
  \_/\_/ \___/|_|  \__,_|_____| .__/ \__,_|_|
                              |_|

      Credit: Madhurend ( m14r41 )
      
""")

def convert_word_to_pdf(word_file, pdf_file):
    try:
        # Initialize Word application
        word = win32com.client.Dispatch("Word.Application")
        # Open the Word document
        doc = word.Documents.Open(word_file)
        # Save the Word document as PDF (FileFormat=17 corresponds to PDF format)
        doc.SaveAs(pdf_file, FileFormat=17)
        # Close the document
        doc.Close()
        # Quit Word application
        word.Quit()
        print(Fore.GREEN + f"Converted {word_file} to PDF.")
    except Exception as e:
        print(Fore.RED + f"Failed to convert {word_file} to PDF. Error: {str(e)}")

def convert_all_word_files_in_folder(folder_path):
    # Loop through all files in the folder
    for filename in os.listdir(folder_path):
        # Check if the file is a Word document
        if filename.endswith(".docx") or filename.endswith(".doc"):
            word_file = os.path.join(folder_path, filename)
            pdf_file = os.path.join(folder_path, filename.replace(".docx", ".pdf").replace(".doc", ".pdf"))
            # Convert the Word file to PDF
            convert_word_to_pdf(word_file, pdf_file)

try:
    # Ask the user to enter the folder path
    folder_path = input(Fore.YELLOW + "Please enter the folder path where Word files are located: ")

    # Check if the provided path exists
    if os.path.isdir(folder_path):
        print(Fore.CYAN + "Starting conversion process...")
        convert_all_word_files_in_folder(folder_path)
    else:
        print(Fore.RED + "The provided folder path is not valid. Please check and try again.")

except KeyboardInterrupt:
    print(Fore.RED + "\nProcess interrupted by the user. Exiting...")
