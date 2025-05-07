import csv
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

class MailMergeGenerator:
    

    DEFAULT_DELIMITERS = ('{{', '}}')

    def __init__(self):
        self.template_content = ""
        self.csv_data: List[Dict[str, str]] = []
        self.placeholders: List[str] = []
        self.output_dir = "output"
        self.current_delimiters = self.DEFAULT_DELIMITERS
        self.errors: List[str] = []

    def load_template(self, template_path: str) -> bool:
        try:
            with open(template_path, 'r', encoding='utf-8') as file:
                self.template_content = file.read()
            self._extract_placeholders()
            return True
        except FileNotFoundError:
            self.errors.append(f"Template file not found: {template_path}")
            return False
        except Exception as e:
            self.errors.append(f"Error reading template: {str(e)}")
            return False

    def _extract_placeholders(self):
        start, end = self.current_delimiters
        pattern = re.escape(start) + r'(.*?)' + re.escape(end)
        matches = re.findall(pattern, self.template_content)
        self.placeholders = list(dict.fromkeys(matches)) 

    def set_delimiters(self, start: str, end: str):
        self.current_delimiters = (start, end)
        if self.template_content:
            self._extract_placeholders()

    def load_csv_data(self, csv_path: str) -> bool:
        try:
            with open(csv_path, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                self.csv_data = [row for row in reader]

            if not self.csv_data:
                self.errors.append("CSV file contains no data")
                return False

            return self._validate_csv_headers()
        except FileNotFoundError:
            self.errors.append(f"CSV file not found: {csv_path}")
            return False
        except Exception as e:
            self.errors.append(f"Error reading CSV: {str(e)}")
            return False

    def _validate_csv_headers(self) -> bool:
        missing = [ph for ph in self.placeholders if ph not in self.csv_data[0]]
        if missing:
            self.errors.append(f"CSV missing required columns: {', '.join(missing)}")
            return False
        return True

    def generate_output(self, output_format: str = "separate", filename_field: Optional[str] = None) -> bool:
        if not self._ready_to_generate():
            return False

        os.makedirs(self.output_dir, exist_ok=True)

        if output_format == "combined":
            return self._generate_combined_output()
        return self._generate_individual_files(filename_field)

    def _ready_to_generate(self) -> bool:
        if not self.template_content:
            self.errors.append("No template loaded")
            return False
        if not self.csv_data:
            self.errors.append("No CSV data loaded")
            return False
        return True

    def _generate_individual_files(self, filename_field: Optional[str]) -> bool:
        success_count = 0
        for i, row in enumerate(self.csv_data):
            try:
                content = self._merge_template(row)
                if filename_field and filename_field in row:
                    base_filename = row[filename_field].strip().replace(' ', '_')
                else:
                    base_filename = f"document_{i+1}"

                filename = f"{base_filename}.txt"
                output_path = os.path.join(self.output_dir, filename)

                counter = 1
                while os.path.exists(output_path):
                    filename = f"{base_filename}_{counter}.txt"
                    output_path = os.path.join(self.output_dir, filename)
                    counter += 1

                with open(output_path, 'w', encoding='utf-8') as file:
                    file.write(content)
                success_count += 1
            except Exception as e:
                self.errors.append(f"Error processing row {i+1}: {str(e)}")

        print(f"Successfully generated {success_count} documents")
        return success_count > 0

    def _generate_combined_output(self) -> bool:
        combined_content = []
        success_count = 0
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(self.output_dir, f"combined_output_{timestamp}.txt")

        for i, row in enumerate(self.csv_data):
            try:
                combined_content.append(self._merge_template(row))
                combined_content.append("\n" + "="*80 + "\n")
                success_count += 1
            except Exception as e:
                self.errors.append(f"Error processing row {i+1}: {str(e)}")

        if success_count:
            try:
                with open(output_path, 'w', encoding='utf-8') as file:
                    file.write("\n".join(combined_content))
                print(f"Successfully generated combined document with {success_count} entries")
                return True
            except Exception as e:
                self.errors.append(f"Error writing combined file: {str(e)}")
        return False

    def _merge_template(self, data: Dict[str, str]) -> str:
        content = self.template_content
        start, end = self.current_delimiters
        for placeholder in self.placeholders:
            value = data.get(placeholder, f"{placeholder}_NOT_FOUND")
            content = content.replace(f"{start}{placeholder}{end}", str(value))
        return content

    def get_errors(self) -> List[str]:
        return self.errors

    def clear_errors(self):
        self.errors.clear()

    @staticmethod
    def get_template_example() -> str:
        return ("""Dear {{first_name}} {{last_name}},

We are pleased to inform you about your recent order #{{order_id}}.

Order details:
- Item: {{product_name}}
- Quantity: {{quantity}}
- Total: {{total_amount}}

Please contact us at {{contact_email}} if you have any questions.

Best regards,
{{company_name}}""")

    @staticmethod
    def get_csv_example() -> str:
        return "first_name,last_name,order_id,product_name,quantity,total_amount,contact_email,company_name"

def main():
    print("\n=== Mail Merge Generator ===")
    merger = MailMergeGenerator()

    while True:
        print("""
Main Menu:
1. Load Template
2. Load CSV Data
3. Set Custom Delimiters
4. Generate Output
5. View Examples
6. Exit""")
        choice = input("Enter your choice (1-6): ").strip()

        if choice == "1":
            path = input("Template file path: ").strip()
            if merger.load_template(path):
                print(f"Loaded. Found {len(merger.placeholders)} placeholders.")
            else:
                for err in merger.get_errors(): print(f"- {err}")
                merger.clear_errors()

        elif choice == "2":
            path = input("CSV file path: ").strip()
            if merger.load_csv_data(path):
                print(f"Loaded {len(merger.csv_data)} records.")
            else:
                for err in merger.get_errors(): print(f"- {err}")
                merger.clear_errors()

        elif choice == "3":
            start = input("Start delimiter (default '{{'): ").strip() or '{{'
            end = input("End delimiter (default '}}'): ").strip() or '}}'
            merger.set_delimiters(start, end)
            print(f"Delimiters set to {start}placeholder{end}.")

        elif choice == "4":
            if not merger.template_content or not merger.csv_data:
                print("Load both template and CSV first.")
                continue
            fmt = input("Output: 1=separate files, 2=combined file: ").strip()
            field = None
            if fmt == "1":
                field = input("CSV field for filename (optional): ").strip() or None
            if merger.generate_output("separate" if fmt == "1" else "combined", field):
                print("Generation complete!")
            else:
                for err in merger.get_errors(): print(f"- {err}")
                merger.clear_errors()

        elif choice == "5":
            print("\nTemplate Example:\n" + merger.get_template_example())
            print("\nCSV Header Example:\n" + merger.get_csv_example())

        elif choice == "6":
            print("Goodbye!")
            break

        else:
            print("Invalid choice. Enter 1-6.")

if __name__ == "__main__":
    main()
