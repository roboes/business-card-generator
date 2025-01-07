## Business Card Generator
# Last update: 2025-01-07


"""About: Fill variables (of a given dataset input) into Merge Fields of a Microsoft Word template using Mail Merge library in Python."""


###############
# Initial Setup
###############

# Erase all declared global variables
globals().clear()


# Import packages
import os

from mailmerge import MailMerge

# from natsort import natsorted
# import openpyxl
import pandas as pd


# Settings

## Set working directory
# os.chdir(path=os.path.join(os.path.expanduser('~'), 'Downloads'))

## Copy-on-Write (will be enabled by default in version 3.0)
if pd.__version__ >= '1.5.0' and pd.__version__ < '3.0.0':
    pd.options.mode.copy_on_write = True


###########
# Functions
###########


def split_dataframe(*, df, chunk_size=10):
    """Split DataFrame into chunks of up to 10 rows (adapted from: https://stackoverflow.com/a/28882020/9195104)."""
    chunks = []
    num_chunks = len(df) // chunk_size + 1

    for i in range(num_chunks):
        chunks.append(df[i * chunk_size : (i + 1) * chunk_size])

    for i in range(len(chunks)):
        chunks[i].index = pd.RangeIndex(start=1, stop=len(chunks[i]) + 1, step=1)
        chunks[i] = chunks[i].reset_index(level=0, drop=False)
        chunks[i] = chunks[i].rename(columns={'index': 'merge_field'})
        chunks[i]['merge_field'] = chunks[i]['merge_field'].astype(dtype='str')
        chunks[i]['merge_field'] = chunks[i]['merge_field'].replace(
            to_replace=r'^(.*)$',
            value=r'guest_\1',
            regex=True,
        )

    return chunks


def business_card_generator(
    *,
    df,
    template,
    output_directory,
    file_name='template_output.docx',
):
    # Import template
    document = MailMerge(template)

    # Get set of Merge Fields
    # document_guest_fields = document.get_merge_fields()

    # Sort Merge Fields
    # document_guest_fields = natsorted(document_guest_fields)

    # Split DataFrame
    df = split_dataframe(df=df, chunk_size=10)

    # Create dict object
    df_list = {}

    for i in range(len(df)):
        df_list[i] = dict(zip(df[i]['merge_field'], df[i]['name']))

    # Fill Microsoft Word template file
    document.merge_templates(
        replacements=df_list.values(),
        separator='continuous_section',
    )
    document.write(file=os.path.join(output_directory, file_name))
    document.close()
