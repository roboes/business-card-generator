## Business Card Generator
# Last update: 2023-05-24


###############
# Initial Setup
###############

# Erase all declared global variables
globals().clear()


# Import packages
# from natsort import natsorted
import os

from mailmerge import MailMerge
# import openpyxl
import pandas as pd


# Set working directory
os.chdir(path=os.path.join(os.path.expanduser('~'), 'Downloads'))




###########
# Functions
###########

# Split DataFrame into chunks of up to 10 rows (adapted from: https://stackoverflow.com/a/28882020/9195104)
def split_dataframe(*, df, chunk_size=10):

    chunks = list()
    num_chunks = len(df)//chunk_size + 1

    for i in range(num_chunks):
        chunks.append(df[i*chunk_size:(i + 1)*chunk_size])

    for i in range(len(chunks)):
        chunks[i].index = pd.RangeIndex(start=1, stop=len(chunks[i]) + 1, step=1)
        chunks[i] = chunks[i].reset_index(level=0, drop=False)
        chunks[i] = chunks[i].rename(columns={'index': 'merge_field'})
        chunks[i]['merge_field'] = chunks[i]['merge_field'].astype(dtype='str')
        chunks[i]['merge_field'] = chunks[i]['merge_field'].str.replace(pat=r'^(.*)$', repl=r'guest_\1', regex=True)

    return chunks



# Business card generator
def business_card_generator(*, df, template, output='template_output.docx'):

    # Import template
    document = MailMerge(template)


    # Get set of Merge Fields
    # document_guest_fields = document.get_merge_fields()


    # Sort Merge Fields
    # document_guest_fields = natsorted(document_guest_fields)


    # Split DataFrame
    df = split_dataframe(df=df, chunk_size=10)


    # Create dictionary object
    df_list = {}

    for i in range(len(df)):
        df_list[i] = dict(zip(df[i]['merge_field'], df[i]['name']))

    # Fill Word Template file
    document.merge_templates(replacements=list(df_list.values()), separator='continuous_section')
    document.write(file=output)
    document.close()




#########################
# Business Card Generator
#########################

# Create example DataFrame with names
df = pd.DataFrame(data=[
    ['Tom'],
    ['Jones'],
    ['Krystal'],
    ['Albert'],
    ['Paloma'],
    ['Shania'],
    ['Max'],
    ['Steve'],
    ['Paul'],
    ['Patrick'],
    ['Lucia'],
    ['Rachel'],
    ['Ray'],
    ['Jessica'],
    ['Julianna'],
    ['Lucille'],
    ['Leandro'],
    ['Vincent'],
    ], index=None, columns=['name'], dtype=None)

# # Import Excel file with names
# df = (pd.read_excel(io='Names.xlsx', sheet_name='List', header=0, index_col=None, skiprows=0, skipfooter=0, dtype=None, engine='openpyxl')
#    .filter(items=['name']))

# Rearrange rows
df = df.sort_values(by=['name'], ignore_index=True)


# Fill/populate Merge Fields from a Microsoft Word file (.docx) from a given DataFrame
business_card_generator(df=df, template='wedding_business_card_template.docx', output='wedding_business_card_template_output.docx')
