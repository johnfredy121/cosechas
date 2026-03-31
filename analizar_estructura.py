import json
import openpyxl

class ExcelAnalyzer:
    def __init__(self, file_path):
        self.file_path = file_path
        self.workbook = openpyxl.load_workbook(file_path, data_only=False)
        self.report = {}

    def analyze(self):
        self.report['sheets'] = {}  
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            self.report['sheets'][sheet_name] = {
                'formulas': {},
                'column_properties': {},
                'hidden_columns': [],
                'cell_formats': {}
            }
            for col in sheet.columns:
                col_letter = col[0].column_letter
                col_hidden = sheet.column_dimensions[col_letter].hidden
                if col_hidden:
                    self.report['sheets'][sheet_name]['hidden_columns'].append(col_letter)

                col_data = []
                for cell in col:
                    col_data.append({
                        'value': cell.value,
                        'formula': cell.formula if cell.formula else '',
                        'format': cell.number_format
                    })
                    if cell.formula:
                        self.report['sheets'][sheet_name]['formulas'][cell.coordinate] = cell.formula
                self.report['sheets'][sheet_name]['column_properties'][col_letter] = col_data

    def generate_report(self):
        return json.dumps(self.report, indent=2)

if __name__ == '__main__':
    file_path = 'your_excel_file.xlsx'
    analyzer = ExcelAnalyzer(file_path)
    analyzer.analyze()
    report = analyzer.generate_report()
    print(report)
