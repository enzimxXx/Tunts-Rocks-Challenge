import gspread 
from google.oauth2.service_account import Credentials
import math

# Authorization scopes for accessing Google Sheets and Google Drive
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Load credentials from the service account file
credentials = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(credentials)  # Authorize using the provided credentials

# Open the spreadsheet
spreadsheet_name = "Engenharia de Software - Desafio Enzo Barbosa"
sheet = client.open(spreadsheet_name).sheet1
total_classes = 60

def get_class_data(sheet, start_row, end_row):
    # Extract data in batches
    data_range = f"C{start_row}:F{end_row}"
    cell_list = sheet.range(data_range)

    class_data = []

    # Extract data from the batch
    for i in range(0, len(cell_list), 4):
        try:
            absences, p1, p2, p3 = map(int, [cell_list[i + j].value for j in range(4)])
            class_data.append((absences, p1, p2, p3))
        except ValueError:
            print(f"Warning: Non-numeric value found in row {start_row + i // 4}")

    return class_data

def main():
    print("Starting processing...")

    # Adjust the loop range to include the last 3 students
    for index in range(4, 28, 3):  # Adjust the loop step as needed
        # Get data in batches for 3 students (adjust as needed)
        class_data = get_class_data(sheet, index, index + 2)

        for i, data in enumerate(class_data):
            absences, p1, p2, p3 = data

            # Calculate average and determine student situation
            average = (p1 + p2 + p3) / 3

            if absences > 0.25 * total_classes:
                situation = "Reprovado por Falta" 
            elif average < 50:
                situation = "Reprovado por Nota"
            elif 50 <= average < 70:
                situation = "Exame Final"
            elif average >= 70:
                situation = "Aprovado"

            sheet.update_cell(index + i, 7, situation)
            
            # Update column H with the calculated final approval grade if in 'Exame Final' situation
            if situation == "Exame Final":
                naf = max(0, 70 - average)  # Calculates the grade needed to reach at least 7 in the final average
                final_approval_grade = average + naf
                sheet.update_cell(index + i, 8, math.ceil(naf))
                
                # Add log line
                print(f"Student {index + i - 3}: Situation updated to '{situation}', Final Approval Grade: {math.ceil(naf)}")
            else:
                sheet.update_cell(index + i, 8, 0)
                print(f"Student {index + i - 3}: Student already passed")

    print("Processing completed.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")
