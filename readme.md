<meta name='keywords' content='Microsoft Word, Merge Fields, Business Card, Wedding Business Card, MailMerge, python'>

# Business Card Generator

## Description

This tool aims to fill variables (of a given dataset input) into Merge Fields of a Microsoft Word template using Mail Merge library in Python. The main features are:

- Fill/populate Merge Fields from a Microsoft Word file (.docx) from a given dataset (an example is provided in the [business-card-generator.py](business-card-generator.py) code).
- In case of more than 10 rows are contained in the input dataset, the template is replicated into multiple pages.

## Output

Two Microsoft Word (.docx) templates are provided:

1. [Simple Business Card Template](templates/simple_business_card_template.docx)

<p align="center">
<img src="./media/simple-business-card-template-output.png" alt="Output" width=510 high=720>
</p>

2. [Wedding Business Card Template](templates/wedding_business_card_template.docx) (required font: [Angella White Font](https://www.dafont.com/angella-white.font)).

<p align="center">
<img src="./media/wedding-business-card-template-output.png" alt="Output" width=510 high=720>
</p>

# Usage

## Python dependencies

```.ps1
python -m pip install docx-mailmerge openpyxl pandas
```

## Functions

### `business_card_generator`

```.py
business_card_generator(df, template, output_name)
```

#### Description

- Fill variables (of a given dataset input) into Merge Fields of a Microsoft Word template.

#### Parameters

- `df`: _DataFrame_. The DataFrame should include a _name_ column, with the names that will be inserted to the Word template file.
- `template`: _str, path object or file-like object_. Word template input file.
- `output_name`: _str, path object or file-like object_. Output of the transformed Word template input file.

## Code Workflow Example

```.py
# Create example DataFrame with names
df = pd.DataFrame(
    data=[
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
    ],
    index=None,
    columns=['name'],
    dtype=None,
)

# # Import Excel file with names
# df = (pd.read_excel(io='Names.xlsx', sheet_name='List', header=0, index_col=None, skiprows=0, skipfooter=0, dtype=None, engine='openpyxl')
#    .filter(items=['name']))

# Rearrange rows
df = df.sort_values(by=['name'], ignore_index=True)


# Fill/populate Merge Fields from a Microsoft Word file (.docx) from a given DataFrame
business_card_generator(
    df=df,
    template=os.path.join(
        os.path.expanduser('~'),
        'Downloads',
        'business-card-generator',
        'templates',
        'wedding_business_card_template.docx',
    ),
    output_directory=os.path.join(os.path.expanduser('~'), 'Downloads'),
    file_name='wedding_business_card_template_output.docx',
)
```
