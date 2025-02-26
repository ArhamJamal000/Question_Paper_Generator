import random
import docx
import math
from question_paper_saver import save_question_paper
from constants import options

class QuestionPaperGenerator:
    def __init__(self, file_path, option, saq_num=None, laq_num=None):
        self.units = self.read_questions_from_word(file_path)
        self.option = option
        self.saq_num = saq_num
        self.laq_num = laq_num

    def read_questions_from_word(self, file_path):
        document = docx.Document(file_path)
        units = {
            "unit_1_saqs": [], "unit_1_laqs": [],
            "unit_2_saqs": [], "unit_2_laqs": [],
            "unit_3_saqs": [], "unit_3_laqs": [],
            "unit_4_saqs": [], "unit_4_laqs": [],
            "unit_5_saqs": [], "unit_5_laqs": []
        }

        for table in document.tables:
            header = table.cell(0, 0).text.strip().lower()
            if "unit" in header:
                unit_number = header.split(" ")[1]
                current_unit = f"unit_{unit_number}"
            else:
                continue

            for row in table.rows[1:]:
                question = row.cells[0].text.strip()
                marks = int(row.cells[1].text.strip())
                co_1 = row.cells[2].text.strip()  # CO
                co_btl = row.cells[3].text.strip()  # BTL

                # Extract numeric values from CO and BTL
                co = co_1.split('-')[1] if 'CO-' in co_1 else co_1.split(' ')[-1] if 'CO' in co_1 else 'N/A'
                btl = co_btl.split('-')[1] if 'BTL-' in co_btl else co_btl.split(' ')[-1] if 'BTL' in co_btl else 'N/A'

                question_info = {"question": question, "marks": marks, "co": co, "btl": btl}

                if marks == 2 and btl in ['1', '2']:
                    units[f"{current_unit}_saqs"].append(question_info)
                elif marks in [6, 12] and btl in ['2', '3', '4', '5', '6']:
                    units[f"{current_unit}_laqs"].append(question_info)

        return units

    import random

    def generate_question_paper(self, selected_units=None, exam_type=None):
        selected_saqs = []
        selected_laqs = []

        if selected_units is None:
            selected_units = range(1, 6)  # Default to Units 1 to 5 if none are passed

        try:
            print(f"Generating question paper for exam type: {exam_type}")

            # Define the exam type (SEE or CIE)
            if exam_type == "SEE":
                total_laqs_needed = 6
                total_saqs_needed = 6
            elif exam_type == "CIE":
                total_laqs_needed = 3
                total_saqs_needed = 3
            else:
                raise ValueError("Invalid exam type. Use 'SEE' or 'CIE'.")

            # Logic for selecting SAQs
            saq_count = 0
            for unit in selected_units:
                saqs_key = f"unit_{unit}_saqs"
                if self.units.get(saqs_key):
                    saq_candidates = self.units[saqs_key]
                    if saq_candidates:
                        unit_saq = random.choice(saq_candidates)
                        selected_saqs.append(unit_saq)
                        saq_count += 1
                        if saq_count >= total_saqs_needed:
                            break

            # Add additional SAQs if needed
            if saq_count < total_saqs_needed:
                for unit in selected_units:
                    saqs_key = f"unit_{unit}_saqs"
                    if self.units.get(saqs_key):
                        saq_candidates = self.units[saqs_key]
                        while saq_candidates and saq_count < total_saqs_needed:
                            selected_saq = random.choice(saq_candidates)
                            selected_saqs.append(selected_saq)
                            saq_candidates.remove(selected_saq)
                            saq_count += 1

            # Logic for selecting LAQs
            laq_count = 0
            remaining_laqs_needed = total_laqs_needed
            laq_6m = []
            laq_12m = []
            btl_2_laqs = []
            btl_3_6_laqs = []

            # Categorize the LAQs based on marks and BTL level
            for unit in selected_units:
                laqs_key = f"unit_{unit}_laqs"
                if self.units.get(laqs_key):
                    laq_candidates = self.units[laqs_key]
                    for laq in laq_candidates:
                        if laq['marks'] == 6:
                            laq_6m.append(laq)
                        elif laq['marks'] == 12:
                            laq_12m.append(laq)
                        elif laq['btl'] == '2':
                            btl_2_laqs.append(laq)
                        elif laq['btl'] in ['3', '4', '5', '6']:
                            btl_3_6_laqs.append(laq)

            # Select exactly one BTL2 question for LAQ
            btl2_laq_selected = False  # Ensure only one BTL2 question is selected
            if btl_2_laqs:
                selected_laq = random.choice(btl_2_laqs)
                selected_laqs.append(selected_laq)
                laq_count += 1
                remaining_laqs_needed -= 1
                btl_2_laqs.remove(selected_laq)
                btl2_laq_selected = True

            # Select remaining LAQs (preferably 12M first, then pair 6M)
            while remaining_laqs_needed > 0:
                if laq_12m and remaining_laqs_needed > 0:
                    selected_laq = random.choice(laq_12m)
                    selected_laqs.append(selected_laq)
                    laq_count += 1
                    remaining_laqs_needed -= 1
                    laq_12m.remove(selected_laq)
                elif len(laq_6m) >= 2 and remaining_laqs_needed > 0:
                    selected_pair = random.sample(laq_6m, 2)
                    paired_laq = {
                        'text': f"Pair of questions: {selected_pair[0]['text']} and {selected_pair[1]['text']}",
                        'marks': selected_pair[0]['marks'] + selected_pair[1]['marks'],
                        'btl': '3-6',  # Represent the combined BTL
                        'unit': f"Units: {selected_pair[0]['unit']} & {selected_pair[1]['unit']}"
                    }
                    selected_laqs.append(paired_laq)
                    laq_count += 1
                    remaining_laqs_needed -= 1
                    laq_6m.remove(selected_pair[0])
                    laq_6m.remove(selected_pair[1])
                elif btl_3_6_laqs and remaining_laqs_needed > 0:
                    selected_laq = random.choice(btl_3_6_laqs)
                    selected_laqs.append(selected_laq)
                    laq_count += 1
                    remaining_laqs_needed -= 1
                    btl_3_6_laqs.remove(selected_laq)

            # Final validation for selected questions
            if len(selected_saqs) != total_saqs_needed or len(selected_laqs) != total_laqs_needed:
                print(f"Warning: Total number of questions selected does not meet the required count.")
                print(f"SAQs selected: {len(selected_saqs)}, LAQs selected: {len(selected_laqs)}")

            print(f"Selected SAQs: {selected_saqs}")
            print(f"Selected LAQs: {selected_laqs}")

            # Return SAQs and LAQs
            return selected_saqs, selected_laqs

        except Exception as e:
            print(f"Error during question paper generation: {e}")
            return [], []  # Always return empty lists in case of error to avoid NoneType

    def save_question_paper(self, file_path, course_name, course_code, examination_date):
        print("Entering save_question_paper method...")
        try:
            saqs, laqs = self.generate_question_paper()
            print("Generated question paper successfully.")

            print("Selected SAQs with CO and BTL values:")
            for saq in saqs:
                print(f"Question: {saq['question']}, CO: {saq['co']}, BTL: {saq['btl']}", flush=True)

            print("Selected LAQs with CO and BTL values:")
            for laq in laqs:
                print(f"Question: {laq['question']}, CO: {laq['co']}, BTL: {laq['btl']}", flush=True)

            # Call the save function
            save_question_paper(saqs, laqs, file_path, course_name, course_code, examination_date, self.option)
        except Exception as e:
            print(f"An error occurred: {e}")