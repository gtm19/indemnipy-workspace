# --8<-- [start:import_setup]
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import logfire
from indemnipy_ai.capabilities.excel import (
    ExcelCapability,
    ExcelRuntimeState,
)
from pydantic import BaseModel, Field
from pydantic_ai import Agent

# Optional - setup Pydantic Logfire for observability
logfire.configure()
logfire.instrument_system_metrics()
logfire.instrument_pydantic_ai()
# --8<-- [end:import_setup]


# --8<-- [start:model_definition]
# Define our Agent dependency model - including a slot for our
# Excel capability runtime state
@dataclass
class Deps:
    excel_runtime_state: ExcelRuntimeState = field(default_factory=ExcelRuntimeState)


# Let's define some Pydantic BaseModels to structure our output
class BrokerInfo(BaseModel):
    contact_name: str
    company: str
    telephone: str | None
    email: str | None


class AddressInfo(BaseModel):
    address_line_1: str
    address_line_2: str | None
    county: str
    postal_code: str
    country_iso2: str  # (1)!


class InsuredInfo(BaseModel):
    name: str
    address: AddressInfo
    n_employees: int
    turnover_usd: list[dict[int, float]] = Field(
        description="Turnover value(s). Each entry is a dict of year: turnover"
    )


class Location(BaseModel):
    id: str | None
    address: str
    country_iso2: str
    average_stock_usd: float
    maxmimum_stock_usd: float


class SubmissionSummary(BaseModel):
    insured: InsuredInfo
    additional_insureds: list[str]
    inception_date: date
    expiry_date: date
    broker_info: BrokerInfo
    submission_date: date
    controlled_drugs: bool = Field(
        description="Does insured handle Schedule 1 or 2 controlled drugs?"
    )
    annual_sendings_value_usd: float = Field(
        description="Annual sendings value in USD for most recent full year"
    )
    locations: list[Location]
    agent_comment: str = Field(
        description="General comments from agent on data extraction. Focus on any uncertainty in the responses"
    )


# --8<-- [end:model_definition]


# --8<-- [start:agent_definition]
# Define our agent. Am using Pydantic AI Gateway but this too is optional
agent = Agent(
    "gateway/openai:gpt-5.6-terra",
    capabilities=[
        ExcelCapability(),  # (1)!
    ],
    output_type=SubmissionSummary,
    name="test-submission-agent",
    deps_type=Deps,
    instructions="Your job is to explore the Excel file(s) and extract the key data for the submission.",
)
# --8<-- [end:agent_definition]


# --8<-- [start:main_run]
def main():
    deps = Deps(
        excel_runtime_state=ExcelRuntimeState(
            spreadsheets=[
                Path(
                    "packages/indemnipy-ai/tests/data/Vantris_Pharmaceuticals_STP_Submission_2026.xlsx"
                )
            ]
        ),
    )
    result = agent.run_sync(
        deps=deps,
    )
    print(result.output)


if __name__ == "__main__":
    main()
# --8<-- [end:main_run]
