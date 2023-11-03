from csv import DictWriter
import glob
import pprint
import re

from docx import Document

ignore = [
    r"This report was prepared on",
    r"Trademark Application / Registration Summary Report"
]

path = "docx/Trade*.docx"
csv_filename = "kls_files.csv"


def is_not_ignore(st:str):
    "check if a string should be ignored"
    status = True
    for ig in ignore:
        if st.find(ig) >=0:
            status = False
    return status


def starts_with_field_name(st, field_delim):
    """tries to find the field name at the start of the paragraph
    Returns True if found
    """
    return par.text.find(field_delim) >= 0


def docx_to_dict(filename: str, field_delim=":"):
    """This reads a docx file data and returns a Dict

    The docx file is assumed to lines of <fieldname> <field_delim> <data>.
    A Dict is returned keyed by the <fieldname>
    """
    d = {}
    curr_key = None
    lines = []

    doc = Document(filename)
    for p in doc.paragraphs:
        text = p.text
        # on a transition to a different field
        if is_not_ignore(text) and text.find(field_delim)>=0:
            # create a key data pair from all the previous lines
            d[curr_key] = "\n".join([lin.strip() for lin in lines])
            print(curr_key,":", d[curr_key])
            # and set up for following lines of data
            lines = []
            chunks = text.split(field_delim)
            curr_key = chunks[0].strip()
            lines.append(field_delim.join(chunks[1:]))
        else:
            lines.append(text.strip())
    # handle last remaining key
    d[curr_key] = "\n".join([lin.strip() for lin in lines])
    del d[None] # remove dummy key
    return d


def scan_folder(path: str, csv_filename: str):
    filenames = glob.glob(path)
    file_data = []
    for f in filenames:
        if f.startswith("~"):
            print("skipping temp file", f)
        else:
            print("reading", f)
            file_data.append(docx_to_dict(f))
    # aggregate all the key fields
    all_keys = {}
    for d in file_data:
        for k in d.keys():
            all_keys[k] = None
    # save to a csv file
    csv_file = open(csv_filename, "w")
    writer = DictWriter(csv_file, all_keys.keys())
    writer.writeheader()
    for d in file_data:
        writer.writerow(d)
    csv_file.close()


if __name__ == "__main__":

    scan_folder(path, csv_filename)

