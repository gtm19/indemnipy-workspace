---
icon: lucide/hammer
---

# Use Cases

This page provides various insurancey examples of how `indemnipy-ai` can be used
alongside [`pydantic-ai`](https://pydantic.dev/docs/ai/overview/) to extract
structured information from spreadsheets.

## Example File

The following code examples can all be run as-is, using any spreadsheet. If you
need an example spreadsheet to test with, you can download the following file:

[Vantris Pharmaceuticals STP Submission 2026.xlsx](
https://github.com/gtm19/indemnipy-workspace/raw/refs/heads/main/packages/
indemnipy-ai/tests/data/Vantris_Pharmaceuticals_STP_Submission_2026.xlsx
)

This is an example of a Cargo Stock Throughput (STP) submission for a
pharmaceutical company.

!!! warning "Disclaimer"

    The example file is a submission generated using the following prompt in Claude:

    ??? quote "Click to expand"

        I need you to create a .xlsx spreadsheet for me.

        This spreadsheet is a simulated attachment to an insurance submission from a broker to a property & casualty insurance company underwriter.

        Please invent details of a pharmaceutical company who want to buy cargo + stock throughput insurance. The spreadsheet should include:

        * An information tab with no real "data" table on it but more human description of the insured, their operations, and the cover they want
        * A "claims" tab which is essentially a table of claims, but:
            * The table is not a named range, or "formatted as a table" in Excel
            * The column headers are formatted for human readability rather than computer readability
            * There is some metadata at the top (say report dates) which sits above the table meaning that the table does not start in row 1
        * A "locations" tab detailing the 30 or so locations where the insured has stock, and the value
        * A "sendings" tab detailing all stock movements: values, description, origin, and destination
        * The above two tables should be named "table" ranges and have no pointless stuff above them on the tab
        * An embedded document (say a pdf of a fake email relating to a fake claim)
        * Some VBA code (even just a hello world in a module)

    The data is therefore entirely fictional - any resemblance to real companies, products, or people is purely coincidental.

## Setup

First, a little setup:

```python
--8<-- "examples/01_key_submission_details.py:import_setup"
```

## Use Case 1: Submission Key Details

If you just want to extract key information with which to populate a submission
dashboard, or even to feed preliminary information into a pricing, the following
is a good starting point.

### Dependencies and Model Definition

Since we want structured output here, and are using Pydantic AI, we can define a
`BaseModel` which describes the information we want to extract.

All of the details of these models are automatically passed to the agent, so it can use them to
guide its extraction of information from the spreadsheet. It is recommended to add more information
to descriptions of the fields, so that the agent has a better understanding of what is expected - but
I am keeping it simple here for brevity.

We also define our `Deps` class here, where the agent will store the spreadsheet and
any tables it finds / creates.

```python
--8<-- "examples/01_key_submission_details.py:model_definition"
```

1.  :earth_americas: Could use [Pydantic Extra Types here](https://pydantic.dev/docs/validation/dev/api/pydantic-extra-types/pydantic_extra_types_country/) but I am being lazy

### Agent Definition

Next, we define our agent. Only things worth noting here are the `ExcelCapability` and
the output type, which is the `SubmissionSummary` model we defined above.

```python
--8<-- "examples/01_key_submission_details.py:agent_definition"
```

1.  :bar_chart: Here is where we drop our `ExcelCapability` into the agent, so it can
    read / manipulate the spreadsheet.

### Running the Agent

Finally, we can run the agent. We pass it the path to the spreadsheet, and it will
return a structured output of the key submission details.

```python
--8<-- "examples/01_key_submission_details.py:main_run"
```

### Output

Here is the output from the agent, which is a structured representation of the key submission details.

The nice thing about using Pydantic is that we end up with a validated instance of our `SubmissionSummary` model, which we can then use in our application: whether that is writing to a database to display in a dashboard, or as the body of an API request to populate information in a pricing model.

You can read the full trace on Pydantic Logfire [here](https://logfire-us.pydantic.dev/public-trace/e0396468-25e9-45b2-bfe2-2a10e6e18915?spanId=25b372f358ef1792).

```json
{
    "insured": {
        "name": "Vantris Pharmaceuticals Ltd",
        "address": {
            "address_line_1": "Vantris House, Chineham Park",
            "address_line_2": "Basingstoke, Hampshire",
            "county": "Hampshire",
            "postal_code": "RG24 8AG",
            "country_iso2": "GB"
        },
        "n_employees": 3050,
        "turnover_usd": [
            {
                "2026": 1140000000
            },
            {
                "2025": 986000000
            }
        ]
    },
    "additional_insureds": [
        "Vantris Sterile Ireland Ltd",
        "Vantris Benelux B.V.",
        "Vantris Pharma Inc. (Delaware)",
        "Vantris do Brasil Ltda."
    ],
    "inception_date": "2026-10-01",
    "expiry_date": "2027-10-01",
    "broker_info": {
        "contact_name": "Jomo Okonjo",
        "company": "Latterworth Marsden Ltd",
        "telephone": "+44 20 7946 0338",
        "email": "j.okonjo@latterworthmarsden.co.uk"
    },
    "submission_date": "2026-08-04",
    "controlled_drugs": false,
    "annual_sendings_value_usd": 1285000000,
    "locations": [
        {
            "id": "LOC-001",
            "address": "Vantris House, Chineham Park, Basingstoke",
            "country_iso2": "GB",
            "average_stock_usd": 28600000,
            "maxmimum_stock_usd": 42800000
        },
        {
            "id": "LOC-002",
            "address": "Unit 4, Wrexham Industrial Estate, Wrexham",
            "country_iso2": "GB",
            "average_stock_usd": 39200000,
            "maxmimum_stock_usd": 61400000
        },
        {
            "id": "LOC-003",
            "address": "Little Island Industrial Park, Cork",
            "country_iso2": "IE",
            "average_stock_usd": 24100000,
            "maxmimum_stock_usd": 38900000
        },
        {
            "id": "LOC-004",
            "address": "Ravensbos 22, Trade Port Noord, Venlo",
            "country_iso2": "NL",
            "average_stock_usd": 21800000,
            "maxmimum_stock_usd": 33500000
        },
        {
            "id": "LOC-005",
            "address": "3820 Airways Boulevard, Memphis",
            "country_iso2": "US",
            "average_stock_usd": 18900000,
            "maxmimum_stock_usd": 29700000
        },
        {
            "id": "LOC-006",
            "address": "Rodovia Presidente Dutra km 225, Guarulhos",
            "country_iso2": "BR",
            "average_stock_usd": 7900000,
            "maxmimum_stock_usd": 12400000
        },
        {
            "id": "LOC-007",
            "address": "Plot 44, Genome Valley, Shameerpet, Hyderabad",
            "country_iso2": "IN",
            "average_stock_usd": 9800000,
            "maxmimum_stock_usd": 14600000
        },
        {
            "id": "LOC-008",
            "address": "Dock Road, Port of Felixstowe, Felixstowe",
            "country_iso2": "GB",
            "average_stock_usd": 4200000,
            "maxmimum_stock_usd": 8900000
        },
        {
            "id": "LOC-009",
            "address": "31 Airport Boulevard, Changi, Singapore",
            "country_iso2": "SG",
            "average_stock_usd": 11400000,
            "maxmimum_stock_usd": 18200000
        },
        {
            "id": "3PL-01",
            "address": "Cargo City Sud, Gebaude 553, Frankfurt",
            "country_iso2": "DE",
            "average_stock_usd": 2900000,
            "maxmimum_stock_usd": 7800000
        },
        {
            "id": "3PL-02",
            "address": "1150 Devon Avenue, Elk Grove Village",
            "country_iso2": "US",
            "average_stock_usd": 2200000,
            "maxmimum_stock_usd": 6400000
        },
        {
            "id": "3PL-03",
            "address": "Distripark Maasvlakte, Aziehaven 12, Rotterdam",
            "country_iso2": "NL",
            "average_stock_usd": 5100000,
            "maxmimum_stock_usd": 9200000
        },
        {
            "id": "3PL-04",
            "address": "Celsiusweg 8, Venlo",
            "country_iso2": "NL",
            "average_stock_usd": 3300000,
            "maxmimum_stock_usd": 5600000
        },
        {
            "id": "3PL-05",
            "address": "Unit 7, Horizon Logistics Park, Dublin",
            "country_iso2": "IE",
            "average_stock_usd": 6700000,
            "maxmimum_stock_usd": 11800000
        },
        {
            "id": "3PL-06",
            "address": "500 Cabot Boulevard, Langhorne",
            "country_iso2": "US",
            "average_stock_usd": 8100000,
            "maxmimum_stock_usd": 13900000
        },
        {
            "id": "3PL-07",
            "address": "3-1-1 Heiwajima, Ota-ku, Tokyo",
            "country_iso2": "JP",
            "average_stock_usd": 5900000,
            "maxmimum_stock_usd": 10200000
        },
        {
            "id": "3PL-08",
            "address": "2-14 Nanko-kita, Suminoe-ku, Osaka",
            "country_iso2": "JP",
            "average_stock_usd": 2400000,
            "maxmimum_stock_usd": 4300000
        },
        {
            "id": "3PL-09",
            "address": "Poligono Industrial Cabanillas, Nave 14, Guadalajara",
            "country_iso2": "ES",
            "average_stock_usd": 3800000,
            "maxmimum_stock_usd": 6900000
        },
        {
            "id": "3PL-10",
            "address": "12 Rue de l'Industrie, Lyon",
            "country_iso2": "FR",
            "average_stock_usd": 4100000,
            "maxmimum_stock_usd": 7400000
        },
        {
            "id": "3PL-11",
            "address": "Via Enrico Fermi 44, Verona",
            "country_iso2": "IT",
            "average_stock_usd": 1700000,
            "maxmimum_stock_usd": 3100000
        },
        {
            "id": "3PL-12",
            "address": "ul. Poznanska 88, Warsaw",
            "country_iso2": "PL",
            "average_stock_usd": 2600000,
            "maxmimum_stock_usd": 4800000
        },
        {
            "id": "3PL-13",
            "address": "Flygfraktvagen 6, Arlanda, Stockholm",
            "country_iso2": "SE",
            "average_stock_usd": 2100000,
            "maxmimum_stock_usd": 3900000
        },
        {
            "id": "3PL-14",
            "address": "Ikitelli OSB, Blok 22, Istanbul",
            "country_iso2": "TR",
            "average_stock_usd": 1400000,
            "maxmimum_stock_usd": 2600000
        },
        {
            "id": "3PL-15",
            "address": "Jebel Ali Free Zone, Plot S30115, Dubai",
            "country_iso2": "AE",
            "average_stock_usd": 4900000,
            "maxmimum_stock_usd": 8600000
        },
        {
            "id": "3PL-16",
            "address": "17 Director Road, Spartan, Johannesburg",
            "country_iso2": "ZA",
            "average_stock_usd": 1900000,
            "maxmimum_stock_usd": 3400000
        },
        {
            "id": "3PL-17",
            "address": "Avenida Papa Joao XXIII 1200, Maua",
            "country_iso2": "BR",
            "average_stock_usd": 2800000,
            "maxmimum_stock_usd": 5200000
        },
        {
            "id": "3PL-18",
            "address": "Camino a Melipilla 14200, Santiago",
            "country_iso2": "CL",
            "average_stock_usd": 1600000,
            "maxmimum_stock_usd": 2900000
        },
        {
            "id": "3PL-19",
            "address": "Survey 118, Medchal Road, Hyderabad",
            "country_iso2": "IN",
            "average_stock_usd": 2200000,
            "maxmimum_stock_usd": 3800000
        },
        {
            "id": "3PL-20",
            "address": "5 Changi North Way, Singapore",
            "country_iso2": "SG",
            "average_stock_usd": 3400000,
            "maxmimum_stock_usd": 6100000
        },
        {
            "id": "3PL-21",
            "address": "22 Kellaway Place, Sydney",
            "country_iso2": "AU",
            "average_stock_usd": 2700000,
            "maxmimum_stock_usd": 4700000
        },
        {
            "id": "3PL-22",
            "address": "88 Gimpo-daero, Gangseo-gu, Seoul",
            "country_iso2": "KR",
            "average_stock_usd": 1300000,
            "maxmimum_stock_usd": 2200000
        },
        {
            "id": "3PL-23",
            "address": "6100 Tomken Road, Mississauga",
            "country_iso2": "CA",
            "average_stock_usd": 3000000,
            "maxmimum_stock_usd": 5300000
        }
    ],
    "agent_comment": "Submission date was parsed from the workbook's text value ('4 August 2026'). The inception date is explicitly stated as 1 October 2026 for a twelve-month policy, so expiry has been recorded as 1 October 2027. Annual sendings of USD 1.285bn is the stated estimate for the 2026/27 period; the submission separately cites USD 1.142bn declared for the expiring period. The 120-row Sendings extract totals USD 93.1046m and is described as representative rather than the full annual register. Locations reconcile to 32 sites, USD 242.0m average stock and USD 400.5m maximum stock."
}
```
