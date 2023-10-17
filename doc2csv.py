from csv import DictWriter
import glob
import pprint

from docx import Document


def starts_with_field_name(par, field_delim):
    """trys to find the field name at the start of the paragraph
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
        # on a transition to a different field
        if starts_with_field_name(p, field_delim):
            # create a key data pair from all the lines
            d[curr_key] = "\n".join([lin.strip() for lin in lines])
            # and set up for following lines of data
            lines = []
            chunks = p.text.split(field_delim)
            curr_key = chunks[0].strip()
            lines.append(field_delim.join(chunks[1:]))
        else:
            lines.append(p.text.strip())
    # handle last remaining key
    d[curr_key] = "\n".join([lin.strip() for lin in lines])
    del d[None]
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


path = "docx/*.docx"
csv_filename = "kls_files.csv"


scan_folder(path, csv_filename)

