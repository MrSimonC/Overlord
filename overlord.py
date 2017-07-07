from custom_modules.mssql import QueryDB
from custom_modules.xlsx import XlsxTools
from custom_modules.outlook import Outlook
import os
import sys
import tempfile
__version__ = '0.4'


class Overlord:
    """
    Looks in an input_folder of sql scripts, and will email Excel result to recipients in "to.txt" found in same folder
    v0.1: 14/Mar/16, made into a class 15/Mar/16
    v0.2: 22/Mar/16: moved to a folder input, with optional dictionary
    v0.3: 30/Mar/16: added in "which reports" were run in to body of email & cleaned up additional_reports
    """
    def __init__(self, input_folder):
        server = 'SERVER'
        database = 'DATABASE'
        username = 'USERNAME'
        password = 'PASSWORD'
        self.db = QueryDB(server, database, username, password)
        self.xlsx_output_path = os.path.join(tempfile.gettempdir(), 'Lorenzo Auto DQ.xlsx')
        self.created_by = []
        self.modified_by = []
        self.xlsx = None
        self.names_paths = self.get_names_paths_from_folder(input_folder)
        self.reports_names_paths = self._set_reports_names_paths(self.names_paths)
        self.to = 'recipient@email.com'
        for x in self.names_paths:
            if x['name'] == 'to':
                self.to = open(x['path']).read()

    def process_all_reports(self, additional_reports=None):
        """
        Process all reports in folder plus more with the additional reports list
        :param additional_reports: [{'name': '', 'path': ''}]
        """
        self.xlsx = None
        self.reports_names_paths += self._set_reports_names_paths(additional_reports) if additional_reports else []
        for report in self.reports_names_paths:
            print('SQL: Running ' + report['name'])
            query_result = self._run(report['path'])
            if query_result:
                try:
                    if not self.xlsx:
                        self.xlsx = XlsxTools()
                        self.xlsx.create_document(query_result, report['name'], self.xlsx_output_path)
                    else:
                        self.xlsx.add_work_sheet(query_result, report['name'])
                except PermissionError:
                    print('Close the file for f# sake.')
                # self.set_created_by(query_result)
                self.set_modified_by(query_result)
        if self.xlsx:
            self.email()

    def _run(self, sql_path):
        return self.db.exec_sql(open(sql_path).read())

    def set_created_by(self, data):
        # Analyse who created items
        for row in data:
            if 'Service Point Created by' in row.keys():
                self.created_by.append(row['Service Point Created by'])
        self.created_by = list(set(self.created_by))

    def set_modified_by(self, data):
        # Analyse who modified items
        for row in data:
            if 'Service Point Modified by' in row.keys():
                self.modified_by.append(row['Service Point Modified by'])
            if 'Session modified by' in row.keys():
                self.modified_by.append(row['Session modified by'])
            if 'Modified by' in row.keys():
                self.modified_by.append(row['Modified by'])
        self.modified_by = list(set(self.modified_by))

    def email(self):
        email = Outlook()
        body = 'Please see attached document for Lorenzo DQ.\n'
        body += '\nReports processed are: ' + ', '.join([x['name'] for x in self.reports_names_paths]) + '\n'
        if self.created_by:
            body += 'Created by includes: ' + ', '.join(self.created_by) + '\n'
        if self.modified_by:
            body += 'Modified by includes: ' + ', '.join(self.modified_by) + '\n'
        body += '\nKind Regards,\nOverlord.'
        email.send(True,
                   self.to,
                   'Lorenzo DQ Report',
                   body,
                   attachments=[self.xlsx_output_path])

    @staticmethod
    def _set_reports_names_paths(names_paths):
        report_names_paths = []
        for report in names_paths:
            _, ext = os.path.splitext(report['path'])
            if ext == '.sql':
                report_names_paths.append(report)
        return report_names_paths

    @staticmethod
    def get_names_paths_from_folder(folder):
        """
        Scans folder for files, outputs list of dictionaries with {'name': ... and 'path'...}
        :param folder: input folder
        :return: [{'name': '', 'path': ''}] name is without extension
        """
        names_and_paths = []
        for file in os.listdir(folder):
            if os.path.isfile(os.path.join(folder, file)):
                names_and_paths += [{'name': os.path.splitext(file)[0], 'path': os.path.join(folder, file)}]
        return names_and_paths


if __name__ == '__main__':
    if hasattr(sys, 'frozen'):
        this_module_path = os.path.dirname(sys.executable)
    else:
        this_module_path = os.path.dirname(os.path.realpath(__file__))
    sql_reports = open(os.path.join(this_module_path, 'overlord.txt')).read()
    o = Overlord(sql_reports)
    o.process_all_reports()

# Compilation:
# from custom_modules.compile_helper import CompileHelp
# c = CompileHelp(r'C:\simon_files_compilation_zone\overlord')
# # c.create_env('pyodbc, openpyxl, pypiwin32')
# c.freeze(r'K:\Coding\Python\nbt work\overlord.py', copy_to=r'C:\Program Files\Overlord\overlord.exe', args='-wF')