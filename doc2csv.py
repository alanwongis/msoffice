from csv import DictWriter
import glob
import pprint
import re

from docx import Document

ignore = [
    r"Trademark Application / Registration Summary Report"
]

path = "docx/Trade*.docx"
csv_filename = "kls_files.csv"

valid_fields = []
f = open("valid_field_names.txt", "r")
for line in f:
    valid_fields.append(line.strip())
f.close()


def is_ignore(st:str):
    "check if a string should be ignored"
    for ig in ignore:
        if st.find(ig) >= 0:
            return True
    return False


def starts_with_field_name(st, field_delim):
    """tries to find the field name at the start of the paragraph
    Returns True if found
    """
    if st.find(field_delim) >= 0:
        possible_key = st.split(field_delim)[0].strip()
        return possible_key in valid_fields
    return False


def is_record_end(st):
    return st.find("This report was prepared on") >= 0


def docx_to_dict(filename: str, field_delim=":"):
    """This reads a docx file data and returns a Dict

    The docx file is assumed to lines of <fieldname> <field_delim> <data>.
    A Dict is returned keyed by the <fieldname>
    """
    records = []
    d = {}
    curr_key = None
    lines = []

    doc = Document(filename)
    for p in doc.paragraphs:
        text = p.text
        # on a transition to a different field
        if is_ignore(text):
            pass
        elif is_record_end(text):
            # finish the last field
            d[curr_key] = "\n".join([lin.strip() for lin in lines])
            records.append(d) # save the record
            d = {} # reset for the net record
        elif starts_with_field_name(text, field_delim): # new field
            # create a key data pair from all the previous lines
            d[curr_key] = "\n".join([lin.strip() for lin in lines])
            # and set up for following lines of data
            lines = []
            chunks = text.split(field_delim)
            curr_key = chunks[0].strip()
            lines.append(field_delim.join(chunks[1:]))
        else:
            lines.append(text.strip())
    # handle last remaining key
    d[curr_key] = "\n".join([lin.strip() for lin in lines])
    records.append(d)# save the last record
    return records


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
    ##all_keys = {}
    ##for d in file_data[0]:
    ##    for k in d[0].keys():
    ##        all_keys[k] = None
    # save to a csv file
    ##csv_file = open(csv_filename, "w")
    ##writer = DictWriter(csv_file, all_keys.keys())
    ##writer.writeheader()
    ##for d in file_data:
    ##    writer.writerow(d)
    ##csv_file.close()
    pprint.pprint(file_data)


def get_valid_field_names():
    records = docx_to_dict("docx\\TrademarkRegistrationSummary.docx")
    field_name_counts = {}
    for rec in records:
        for k in rec.keys():
            field_name_counts.setdefault(k,0)
            field_name_counts[k] +=1
    valid_field_names =[]
    for k in field_name_counts.keys():
        if field_name_counts[k] >10:
            valid_field_names.append(k)
    f = open("valid_field_names.txt", "w")
    for k in valid_field_names:
        f.write(k+"\n")
    f.close()

    print(valid_field_names)




if __name__ == "__main__":
    scan_folder(path, csv_filename)
    #get_valid_field_names()
