# Generic Spreadsheet ETL Automation

This is a non-confidential, reusable example of a four-step Excel ETL workflow.
It contains no organization names, cloud URLs, customer identifiers, or
business-system labels. Workbook names, worksheet names, column names, output
labels, and lookup rules are supplied through `config.json` and the local
workbooks.

The original project is not modified by this folder.

## Included structure

```text
Generic Spreadsheet ETL Automation/
├── main.py
├── config.example.json
├── requirements.txt
├── etl_steps/
│   ├── allocation_addition.py
│   ├── charge_addition.py
│   ├── common.py
│   ├── ledger_addition.py
│   └── record_addition.py
├── Required Data/
│   ├── README.md
│   └── Empty Template File/
│       └── README.md
└── Output/
```

The workbook files are intentionally not included. The `Required Data` folder
contains placeholder instructions instead of fabricated `.xlsx` files.

## Setup

1. Copy `config.example.json` to `config.json`.
2. Replace the placeholder workbook paths, worksheet names, column names, and
   template settings in `config.json`.
3. Add the source workbook and mapping workbook under the configured paths.
4. Add one template workbook for each configured ETL step.
5. Use an existing Python installation that already has the packages listed in
   `requirements.txt`.

No Python environment is created by this project.

## Run

From this folder, run:

```text
python main.py
```

For a mapping workbook containing multiple entity codes, either set
`application.entity_code` in `config.json` or select one interactively. A
specific run can also be selected without changing the file:

```text
python main.py --entity SAMPLE_ENTITY
```

Use another configuration file when testing a separate workbook layout:

```text
python main.py --config path\to\another-config.json
```

## Expected input layout

The example configuration uses these generic source columns:

- `Record ID`
- `Entity Code`
- `Reference`
- `Start Date`
- `End Date`
- `Amount`
- `Cost Center`
- `Processing Status`

Change these names in the configuration when the source workbook uses a
different schema. The controller validates configured required columns before
writing an output workbook.

The mapping sheet must contain at least the configured `Entity Code` column and
one row for the selected entity. Other mapping columns are referenced by the
step-specific assignments in the example configuration and may be renamed
there.

## Template expectations

The example configuration defines four generic template workbooks:

1. Record addition
2. Allocation addition
3. Charge addition
4. Ledger addition

Each template must contain the configured worksheet. The allocation template
also contains a lookup worksheet. The ledger template is checked against the
configured header list; edit `expected_headers` when using a different layout.

The step modules preserve template formatting, copy lookup values into the
allocation output, and add list validation to configured allocation columns.
The controller normalizes only generated data rows to Calibri 11 and leaves
header rows and merged cells unchanged.

## Processing behavior

The controller:

1. Reads the source and mapping workbooks once.
2. Trims the configured number of trailing source rows.
3. Applies the optional processing-status filter.
4. Restricts records to the selected entity when enabled.
5. Removes blank or configured total rows from the first source column.
6. Optionally applies a generic active-month date window.
7. Reports incomplete rows before output generation.
8. Reuses the filtered source and selected mapping DataFrames across all steps.
9. Runs the configured steps in `step_order`.
10. Writes a local audit log, run summary, and any validation reports.

No email, cloud-storage, or desktop application integration is enabled. This
keeps the example safe to share and easy to adapt. A separate integration can
consume the generated reports if an organization needs notifications.

## Output

Outputs are written below the configured output directory using the configured
folder and filename templates. The default pattern is:

```text
Output/YYYY/MM. Mon/<entity>/
├── 01_Record_<month>.xlsx
├── 02_Allocation_<month>.xlsx
├── 03_Charge_<month>.xlsx
├── 04_Ledger_<month>.xlsx
├── ETL_Run_Summary.json
└── ETL_Audit_YYYYMMDD_HHMMSS.log
```

If source completeness validation fails, an `Incomplete_Source_Rows_<month>.xlsx`
report is written and the run stops by default. If an allocation lookup cannot
be resolved, a `Lookup_Failures_<month>.xlsx` report is written and no partial
allocation output is presented as complete.

## Adapting the example

- Change `application` paths and worksheet names first.
- Change `processing.required_source_columns` to match the source workbook.
- Update the step-specific source and mapping assignments.
- Update allocation lookup rules and dropdown columns.
- Update ledger `expected_headers` and assignments together.
- Keep values that vary by entity in the mapping workbook rather than in
  Python code.

The step modules are intentionally independent so one can be replaced without
rewriting the controller.
