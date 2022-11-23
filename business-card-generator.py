## Business Card Generator
# Last update: 2022-11-03


#########################
# ---- initial_setup ----
#########################

## Erase all declared global variables
globals().clear()


## Import packages
from natsort import natsorted
import os

from mailmerge import MailMerge
import pandas as pd


## Set working directory to user's 'Downloads' folder
os.chdir(os.path.join(os.path.expanduser('~'), 'Downloads'))




#####################
# ---- functions ----
#####################

## split_dataframe
# Split dataframe into chunks of up to 10 rows
# Adapted from: https://stackoverflow.com/a/28882020/9195104
def split_dataframe(df, chunk_size=10):

    chunks = list()
    num_chunks = len(df) // chunk_size+1

    for i in range(num_chunks):
        chunks.append(df[i*chunk_size:(i+1)*chunk_size])

    for i in range(len(chunks)):
        chunks[i].index = pd.RangeIndex(start=1, stop=len(chunks[i])+1, step=1)
        chunks[i].reset_index(inplace=True, level=0)
        chunks[i] = chunks[i].rename(columns={'index': 'merge_field'})
        chunks[i]['merge_field'] = chunks[i]['merge_field'].astype(str)
        chunks[i]['merge_field'] = chunks[i]['merge_field'].str.replace(r'^(.*)$', r'guest_\1', regex=True)

    return chunks



## business_card_generator
def business_card_generator(df, template, output_name='template_output.docx'):

    # Import template
    document = MailMerge(template)

    # Get set of Merge Fields
    # document_guest_fields = document.get_merge_fields()

    # Sort Merge Fields
    # document_guest_fields = natsorted(document_guest_fields)

    # Split dataframe
    df = split_dataframe(df, chunk_size=10)

    # Create dataframe dictionary object
    df_list = {}

    for i in range(len(df)):
        df_list[i] = dict(zip(df[i]['merge_field'], df[i]['name']))

    # Fill Word Template file
    document.merge_templates(list(df_list.values()), separator='continuous_section')
    document.write(output_name)
    document.close()
    



###################################
# ---- business-card-generator ----
###################################

## Create simple dataframe with names
df = [
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
    ]
df = pd.DataFrame(df, columns = ['name'])

## Import dataframe with names
# df = pd.read_excel('Names.xlsx', sheet_name='List', engine='openpyxl')
# df = df.filter(['name'])

# Rarrange rows
df = df.sort_values(by=['name'], ignore_index=True)


## Fill/populate Merge Fields from a Microsoft Word file (.docx) from a given dataframe
business_card_generator(df=df, template='wedding_business_card_template.docx', output_name='wedding_business_card_template_output.docx')
