import os
import sys
import shutil
from PyPDF2 import PdfFileMerger, PdfFileReader
import openpyxl
import datetime as dt
import traceback

INVOICES_PATH = f"C:{os.path.sep}Users{os.path.sep}Malcolm{os.path.sep}Desktop{os.path.sep}Working Folder{os.path.sep}Buyers Invoices{os.path.sep}"
REGISTERS_PATH = f"C:{os.path.sep}Users{os.path.sep}Malcolm{os.path.sep}Desktop{os.path.sep}Working Folder{os.path.sep}Buyers RR{os.path.sep}"
MATCHED_PATH = f"C:{os.path.sep}Users{os.path.sep}Malcolm{os.path.sep}Desktop{os.path.sep}Working Folder{os.path.sep}Matched{os.path.sep}"


class Matching:

    def __init__(self):
        self.report = MatchingReporting()
        # Declares invoices and registers as a list of dict
        self.invoices = self.get_files(path=INVOICES_PATH)
        self.registers = self.get_files(path=REGISTERS_PATH)

    def run(self):
        for invoice in self.invoices:
            # For each invoice, we search the registers for a match
            self.search_registers_for_match(invoice=invoice)
        for invoice in self.invoices:
            # For each invoice, we search the registers for a variance
            self.search_registers_for_variance(invoice=invoice)

        self.report.send_report()

    def search_registers_for_match(self, invoice):
        for rr in self.registers:
            # Iterates over the registers to find single matches and merges the files.
            if self.is_match(inv=invoice, register=rr):
                # Ensures this is the only match, if it is we will merge, if not we will add to the report for A/P to look into manually.
                if self.only_match(po=invoice['po'], amount=invoice['subtotal']):
                    self.merge_files(invoice=invoice, register=rr)
                else:
                    self.report.append_multi_match(invoice=invoice)
                self.invoices.remove(invoice)
                self.registers.remove(rr)

    def search_registers_for_variance(self, invoice):
        for rr in self.registers:
            # Notifies A/P of variances to ensure it's a match with a variance or not a match at all.
            if self.has_variance(inv=invoice, register=rr):
                self.report.append_variance(invoice=invoice, register=rr)

    def only_match(self, po, amount):
        return self.only_match_in_list(data=self.invoices, po=po, amount=amount) and self.only_match_in_list(data=self.registers, po=po, amount=amount)

    def merge_files(self, invoice, register):
        # Moves the invoice and register match to the matching folder and creates a third file for the merged PDF and saves the two separate documents as backup.
        shutil.move(f"{INVOICES_PATH}{invoice['file_name']}", f"{MATCHED_PATH}{invoice['file_name']}")
        shutil.move(f"{REGISTERS_PATH}{register['file_name']}", f"{MATCHED_PATH}{register['file_name']}")

        merged_file = PdfFileMerger()
        merged_file.append(PdfFileReader(f"{MATCHED_PATH}{invoice['file_name']}"))
        merged_file.append(PdfFileReader(f"{MATCHED_PATH}{register['file_name']}"))
        merged_file.write(f"{MATCHED_PATH}MATCHED-{invoice['number']}.pdf")

        self.report.append_match(invoice=invoice, register=register)

    @staticmethod
    def only_match_in_list(data, po, amount):
        # Searches the list and ensures only one match
        count = 0
        for index in data:
            if count >= 2:
                break
            elif index['po'] == po and index['subtotal'] == amount:
                count += 1
        return count < 2

    @staticmethod
    def is_match(inv, register):
        # compares the PO numbers and the subtotals and determines if there is a match
        return register['po'] == inv['po'] and register['subtotal'] == inv['subtotal']

    @staticmethod
    def has_variance(inv, register):
        # compares the PO numbers and the subtotals and determines if there is a small variance
        return register['po'] == inv['po'] and register['subtotal'] != inv['subtotal'] and abs(float(inv['subtotal']) - float(register['subtotal'])) <= 10

    @staticmethod
    def get_files(path):
        # Goes to the file path passed as a parameter and iterates all of the files and creates the dictionary of data of the file names.
        f = []
        os.chdir(path)
        for file in os.listdir():
            splits = file.split(" ")
            if len(splits) > 2:
                f.append({'file_name': file, 'po': splits[0], "number": splits[1], 'subtotal': splits[2].replace(".pdf", "")})
        return f


class MatchingReporting:
    def __init__(self):
        print("Loading the notification handler...")
        self.matches = []
        self.variances = []
        self.multiple_matches = []
        self.wb = None

    def send_report(self):
        self.wb = openpyxl.load_workbook(f'{MATCHED_PATH}AutoMatchReport.xlsx')
        self.save_matches()
        self.save_variances()
        self.save_multiple_matches()
        self.wb.save(f'{MATCHED_PATH}AutoMatchReport.xlsx')
        self.wb.close()

    def save_matches(self):
        self.wb.active = self.wb['Matches']
        for match in self.matches:
            self.wb.active.append(match)

    def save_variances(self):
        self.wb.active = self.wb['Variances']
        for variance in self.variances:
            self.wb.active.append(variance)

    def save_multiple_matches(self):
        self.wb.active = self.wb['MultipleMatches']
        for match in self.multiple_matches:
            self.wb.active.append(match)

    def append_match(self, invoice, register):
        print(f"Found match - Invoice: {invoice['file_name']} - RR: {register['file_name']}")
        self.matches.append((dt.datetime.today().date(), invoice['number'], register['number']))

    def append_variance(self, invoice, register):
        print(f"Found variance - PO: {invoice['po']} Invoice: {invoice['number']} Amount: {invoice['subtotal']}")
        self.variances.append((dt.datetime.today().date(), invoice['number'], register['number']))

    def append_multi_match(self, invoice):
        print(f"More than one match for PO: {invoice['po']} Invoice: {invoice['number']} Amount: {invoice['subtotal']}")
        self.multiple_matches.append((dt.datetime.today().date(), invoice['po'], invoice['number']))


def print_stack():
    print(traceback.format_exc())
    exception_type, exception_object, exception_traceback = sys.exc_info()
    filename = exception_traceback.tb_frame.f_code.co_filename
    line_number = exception_traceback.tb_lineno
    exception_message = f'Exception Type: {exception_type} - File Name: {filename} - Line Number: {line_number}'
    f = open(f'{MATCHED_PATH}ErrorLog.txt', "a")
    f.write(exception_message)
    f.close()


if __name__ == '__main__':

    try:
        Matching().run()
    except Exception as e:
        print_stack()
