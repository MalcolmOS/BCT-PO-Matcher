import os
import shutil
from PyPDF2 import PdfFileMerger, PdfFileReader
import datetime as dt
import time

INVOICES_PATH = ''
REGISTERS_PATH = ''
MATCHED_PATH = ''


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
        self.report.send_report()

    def search_registers_for_match(self, invoice):
        for rr in self.registers:
            # Iterates over the registers to find single matches and merges the files.
            if self.is_match(inv=invoice, register=rr):
                # Ensures this is the only match, if it is we will merge, if not we will add to the report for A/P to look into manually.
                if self.only_match(po=invoice['po'], amount=invoice['subtotal']):
                    self.merge_files(invoice=invoice['file_name'], register=rr['file_name'])
                else:
                    self.report.append(f"More than one match for PO: {invoice['po']} Invoice: {invoice['number']} Amount: {invoice['subtotal']}")
            elif self.has_variance(inv=invoice, register=rr):
                # Notifies A/P of variances to ensure it's a match with a variance or not a match at all.
                self.report.append(f"Found variance - PO: {invoice['po']} Invoice: {invoice['number']} Amount: {invoice['subtotal']}")

    def only_match(self, po, amount):
        return self.only_invoice_match(po=po, amount=amount) and self.only_register_match(po=po, amount=amount)

    # Counts how many invoices have this PO and subtotal
    def only_invoice_match(self, po, amount):
        count = 0
        for invoice in self.invoices:
            if count >= 2:
                break
            elif invoice['po'] == po and invoice['subtotal'] == amount:
                count += 1
        return count < 2

    # Counts how many registers have this PO and subtotal
    def only_register_match(self, po, amount):
        count = 0
        for rr in self.registers:
            if count >= 2:
                break
            elif rr['po'] == po and rr['subtotal'] == amount:
                count += 1
        return count < 2

    @staticmethod
    def is_match(inv, register):
        # compares the PO numbers and the subtotals and determines if there is a match
        return register['po'] == inv['po'] and register['subtotal'] == inv['subtotal']

    @staticmethod
    def has_variance(inv, register):
        # compares the PO numbers and the subtotals and determines if there is a small variance
        return register['po'] == inv['po'] and abs(float(inv['subtotal']) - float(register['subtotal'])) < 5

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

    @staticmethod
    def merge_files(invoice, register):
        # Moves the invoice and register match to the matching folder and creates a third file for the merged PDF and saves the two separate documents as backup.
        print(f"Found match - Invoice: {invoice} - RR: {register}")
        shutil.move(f"{INVOICES_PATH}{os.path.sep}{invoice}", f"{MATCHED_PATH}{os.path.sep}{invoice}")
        shutil.move(f"{REGISTERS_PATH}{os.path.sep}{register}", f"{MATCHED_PATH}{os.path.sep}{register}")

        merged_file = PdfFileMerger()
        merged_file.append(PdfFileReader(f"{MATCHED_PATH}{os.path.sep}{invoice}"))
        merged_file.append(PdfFileReader(f"{MATCHED_PATH}{os.path.sep}{register}"))
        merged_file.write(f"{MATCHED_PATH}{os.path.sep}MATCHED-{invoice}")


class MatchingReporting:
    def __init__(self):
        print("Loading the notification handler...")
        self.output = ''

    def send_report(self):
        print("Sending report")
        print(self.output)
        self.output = ''

    def append(self, content):
        self.output += f'{content}\n'


if __name__ == '__main__':

    print("Starting...")
    while True:
        try:
            # Best to run as a scheduled runtime and not leave a background process but if we ran it as a background process we can do this.
            # If the hour of the day is midnight, we will run this matching code.
            if dt.datetime.today().hour == 0:
                Matching().run()
        except Exception as e:
            # Prints out any errors should they happen
            print(e)
        finally:
            # sleeps for one hour until re-looping and starting over
            time.sleep(3600)
