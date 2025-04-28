from typing import Callable, List, Dict, Tuple
import pandas as pd
import zipfile
from lxml import etree
import functools as ft
from pydantic import BaseModel, field_validator, Field
import os


# Basic type definitions
Cell = str | float
Row = List[Cell]
Table = List[Row]
Document = List[Table]


# Custom exception classes
class DocumentProcessingError(Exception):
    """Base class for document processing errors."""

    pass


class DocumentLoadingError(DocumentProcessingError):
    """Error when loading a document."""

    pass


class TableProcessingError(DocumentProcessingError):
    """Error when processing a table."""

    pass


class ExportError(DocumentProcessingError):
    """Error when exporting processed data."""

    pass


# Pydantic models for validation
class TransactionRowParsingConfig(BaseModel):
    field_count: int = Field(
        3, gt=0, description="Number of fields to extract from transaction description"
    )
    id_test_func: Callable[[str], bool] = Field(
        default_factory=lambda: lambda s: s.isdigit(),
        description="Function to identify ID field in transaction",
    )

    class Config:
        arbitrary_types_allowed = True


class TableFormat(BaseModel):
    header_len: int = Field(..., ge=0, description="Number of header rows to skip")
    footer_len: int = Field(..., ge=0, description="Number of footer rows to skip")
    account_cell_index: int = Field(
        ..., ge=0, description="Index of account cell in detail row"
    )
    debit_cell_index: int = Field(
        ..., ge=0, description="Index of debit cell in detail row"
    )
    credit_cell_index: int = Field(
        ..., ge=0, description="Index of credit cell in detail row"
    )
    transaction_row_parsing_config: TransactionRowParsingConfig

    class Config:
        arbitrary_types_allowed = True


class InputDocumentFormat(BaseModel):
    path: str
    table_index: int = Field(..., ge=0, description="Index of table to process")

    @field_validator("path")
    def path_must_exist(cls, v: str):
        if not os.path.exists(v):
            raise ValueError(f"Document path '{v}' does not exist")
        if not os.path.isfile(v):
            raise ValueError(f"'{v}' is not a file")
        if not v.endswith(".docx"):
            raise ValueError(f"'{v}' is not a .docx file")
        return v


class OutputDocumentFormat(BaseModel):
    path: str
    columns: List[str] = Field(..., min_length=1)

    @field_validator("path")
    def path_must_be_valid(cls, v: str):
        try:
            parent_dir = os.path.dirname(v) or "."
            if parent_dir and not os.path.exists(parent_dir):
                raise ValueError(f"Directory '{parent_dir}' does not exist")
            if not v.endswith((".xlsx", ".csv")):
                raise ValueError(f"Output file '{v}' must be .xlsx or .csv")
        except Exception as e:
            raise ValueError(f"Invalid output path: {str(e)}")
        return v


class LoadingConfig(BaseModel):
    loading_strategy: Callable[[InputDocumentFormat], Document]
    table_choose_strategy: Callable[[Document, InputDocumentFormat], Table]
    document_format: InputDocumentFormat

    class Config:
        arbitrary_types_allowed = True


class ProcessingConfiguration(BaseModel):
    header_processing_strategy: Callable[[Table, TableFormat], List[Row]]
    footer_processing_strategy: Callable[[Table, TableFormat], List[Row]]
    detail_row_processing_strategy: Callable[[Row, TableFormat], Row]
    transaction_row_processing_strategy: Callable[[Row, TableFormat], Row]
    combine_rows_strategy: Callable[[Row, Row, TableFormat], Row]
    table_format: TableFormat

    class Config:
        arbitrary_types_allowed = True


class ExportConfig(BaseModel):
    export_strategy: Callable[[Table, OutputDocumentFormat], None]
    output_document_format: OutputDocumentFormat

    class Config:
        arbitrary_types_allowed = True


def load_xml_table(table_element: etree._Element, namespaces: Dict[str, str]) -> Table:  # type: ignore
    """
    Extract table data from XML element.

    Args:
        table_element: XML element representing a table
        namespaces: XML namespaces

    Returns:
        Extracted table as list of rows

    Raises:
        TableProcessingError: If XML parsing fails
    """
    try:
        table: Table = []
        row_elements = table_element.xpath(".//w:tr", namespaces=namespaces)

        for row_element in row_elements:
            row: Row = []
            for cell_element in row_element.xpath(".//w:tc", namespaces=namespaces):
                text = " ".join(
                    [
                        t.text or ""  # Handle None values
                        for t in cell_element.xpath(".//w:t", namespaces=namespaces)
                        if t.text and t.text.strip()
                    ]
                )
                row.append(text)
            if row:  # Only add non-empty rows
                table.append(row)
        return table
    except Exception as e:
        raise TableProcessingError(f"Failed to extract table from XML: {str(e)}")


def load_xml_document(input_document_format: InputDocumentFormat) -> Document:
    """
    Load tables from a DOCX document by parsing its XML structure.

    Args:
        input_document_format: Document format configuration

    Returns:
        List of tables extracted from the document

    Raises:
        DocumentLoadingError: If DOCX parsing fails
    """
    docx_path = input_document_format.path
    try:
        with zipfile.ZipFile(docx_path) as zf:
            xml_content = zf.read("word/document.xml")

        # Parse the XML
        root = etree.fromstring(xml_content)
        namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }

        # Find all tables
        table_elements = root.xpath("//w:tbl", namespaces=namespaces)
        if not table_elements:
            raise DocumentLoadingError("No tables found in document")

        tables: Document = []
        for table_element in table_elements:
            table = load_xml_table(table_element, namespaces)
            tables.append(table)
        return tables
    except zipfile.BadZipFile:
        raise DocumentLoadingError(f"'{docx_path}' is not a valid DOCX file")
    except KeyError:
        raise DocumentLoadingError(
            "Document structure error: word/document.xml not found"
        )
    except etree.XMLSyntaxError as e:
        raise DocumentLoadingError(f"XML parsing error: {str(e)}")
    except Exception as e:
        raise DocumentLoadingError(f"Failed to load document: {str(e)}")


def choose_table(tables: Document, input_document_format: InputDocumentFormat) -> Table:
    """
    Select a specific table from the document.

    Args:
        tables: List of tables
        input_document_format: Document format configuration

    Returns:
        Selected table

    Raises:
        TableProcessingError: If table index is out of range
    """
    try:
        if not tables:
            raise TableProcessingError("No tables available in the document")

        index = input_document_format.table_index
        if index < 0 or index >= len(tables):
            raise TableProcessingError(
                f"Table index {index} out of range (0-{len(tables) - 1})"
            )

        table = tables[index]
        if not table:
            raise TableProcessingError(f"Selected table {index} is empty")

        return table
    except TableProcessingError:
        raise
    except Exception as e:
        raise TableProcessingError(f"Error selecting table: {str(e)}")


def empty_header(table: Table, table_format: TableFormat) -> List[Row]:
    """Returns an empty header."""
    return []


def empty_footer(table: Table, table_format: TableFormat) -> List[Row]:
    """Returns an empty footer."""
    return []


def validate_row_index(row: Row, index: int, function_name: str) -> None:
    """
    Validate that a row has enough elements to access the given index.

    Args:
        row: The row to validate
        index: The index to access
        function_name: Name of the calling function for error reporting

    Raises:
        TableProcessingError: If index is out of range
    """
    if not row:
        raise TableProcessingError(f"{function_name}: Row is empty")
    if index < 0 or index >= len(row):
        raise TableProcessingError(
            f"{function_name}: Index {index} out of range (0-{len(row) - 1})"
        )


def process_detail_row_and_process_account(
    row: Row,
    table_format: TableFormat,
    process_func: Callable[[Cell], Cell],
) -> Row:
    """
    Process a detail row by applying a function to the account cell.

    Args:
        row: The detail row
        table_format: Table format configuration
        process_func: Function to apply to the account cell

    Returns:
        Processed row

    Raises:
        TableProcessingError: If row processing fails
    """
    try:
        # Make a copy to avoid modifying the original
        row_copy = row.copy()

        validate_row_index(
            row_copy,
            table_format.account_cell_index,
            "process_detail_row_and_process_account",
        )

        row_copy[table_format.account_cell_index] = process_func(
            row_copy[table_format.account_cell_index]
        )
        return row_copy
    except TableProcessingError:
        raise
    except Exception as e:
        raise TableProcessingError(f"Error processing detail row: {str(e)}")


def process_detail_row_and_process_account_debit_credit(
    row: Row,
    table_format: TableFormat,
    process_account_func: Callable[[Cell], Cell],
    process_debit_func: Callable[[Cell], Cell],
    process_credit_func: Callable[[Cell], Cell],
) -> Row:
    """
    Process a detail row by applying functions to the account, debit, and credit cells.

    Args:
        row: The detail row
        table_format: Table format configuration
        process_account_func: Function to apply to the account cell
        process_debit_func: Function to apply to the debit cell
        process_credit_func: Function to apply to the credit cell

    Returns:
        Processed row

    Raises:
        TableProcessingError: If row processing fails
    """
    try:
        # Make a copy to avoid modifying the original
        row_copy = row.copy()

        validate_row_index(
            row_copy,
            table_format.account_cell_index,
            "process_detail_row_and_process_account_debit_credit",
        )
        validate_row_index(
            row_copy,
            table_format.debit_cell_index,
            "process_detail_row_and_process_account_debit_credit",
        )
        validate_row_index(
            row_copy,
            table_format.credit_cell_index,
            "process_detail_row_and_process_account_debit_credit",
        )

        row_copy[table_format.account_cell_index] = process_account_func(
            row_copy[table_format.account_cell_index]
        )
        row_copy[table_format.debit_cell_index] = process_debit_func(
            row_copy[table_format.debit_cell_index]
        )
        row_copy[table_format.credit_cell_index] = process_credit_func(
            row_copy[table_format.credit_cell_index]
        )

        return row_copy
    except TableProcessingError:
        raise
    except Exception as e:
        raise TableProcessingError(f"Error processing detail row: {str(e)}")


def parse_transaction_description(
    row: Row,
    table_format: TableFormat,
) -> Row:
    """
    Parse transaction description into counterparty, ID, and description.

    Args:
        row: The transaction row
        table_format: Table format configuration

    Returns:
        Parsed transaction parts

    Raises:
        TableProcessingError: If row processing fails
    """
    try:
        config = table_format.transaction_row_parsing_config

        if not row:
            return [""] * config.field_count

        if len(row) == 0 or not row[0]:
            return [""] * config.field_count

        if not isinstance(row[0], str):
            raise TableProcessingError(
                f"Expected string in row[0], got {type(row[0]).__name__}"
            )

        text = row[0].split(" ")
        result: Row = [""] * config.field_count

        # Try to find ID and split fields accordingly
        id_found = False
        for i, word in enumerate(text):
            if config.id_test_func(word):
                result[0] = " ".join(text[:i]).strip()  # Counterparty
                result[1] = word  # ID
                result[2] = " ".join(text[i + 1 :]).strip()  # Description
                id_found = True
                break

        # If no ID found, put everything in the first field
        if not id_found and text:
            result[0] = " ".join(text).strip()

        return result
    except Exception as e:
        raise TableProcessingError(f"Error parsing transaction description: {str(e)}")


def combine_rows(
    detail_row: Row, transaction_row: Row, table_format: TableFormat
) -> Row:
    """
    Combine detail and transaction rows into a single row.

    Args:
        detail_row: Detail row
        transaction_row: Transaction row
        table_format: Table format configuration

    Returns:
        Combined row

    Raises:
        TableProcessingError: If row combining fails
    """
    try:
        if not detail_row:
            raise TableProcessingError("Detail row is empty")
        if not transaction_row:
            raise TableProcessingError("Transaction row is empty")
        return detail_row + transaction_row
    except TableProcessingError:
        raise
    except Exception as e:
        raise TableProcessingError(f"Error combining rows: {str(e)}")


def extract_transactions(config: ProcessingConfiguration, table: Table) -> Table:
    """
    Extract transactions from a table.

    Args:
        config: Processing configuration
        table: Input table

    Returns:
        Processed transactions table

    Raises:
        TableProcessingError: If transaction extraction fails
    """
    try:
        if not table:
            raise TableProcessingError("Table is empty")

        header_len = config.table_format.header_len
        footer_len = config.table_format.footer_len

        # Validate that the table has enough rows
        min_rows = header_len + footer_len
        if len(table) < min_rows:
            raise TableProcessingError(
                f"Table has {len(table)} rows, needs at least {min_rows} rows"
            )

        header = config.header_processing_strategy(table, config.table_format)
        footer = config.footer_processing_strategy(table, config.table_format)

        table_start = header_len
        table_end = len(table) - footer_len

        transactions: Table = []
        i = table_start
        while i < table_end - 1:
            try:
                # Process the detail row
                detail_row = config.detail_row_processing_strategy(
                    table[i], config.table_format
                )

                # Process the transaction row
                transaction_row = config.transaction_row_processing_strategy(
                    table[i + 1], config.table_format
                )

                # Combine the rows
                combined_row = config.combine_rows_strategy(
                    detail_row, transaction_row, config.table_format
                )

                transactions.append(combined_row)
                i += 2  # Move to next transaction pair
            except IndexError:
                raise TableProcessingError(f"Index error at row {i}/{len(table)}")
            except Exception as e:
                raise TableProcessingError(
                    f"Error processing rows {i}-{i + 1}: {str(e)}"
                )

        return header + transactions + footer
    except TableProcessingError:
        raise
    except Exception as e:
        raise TableProcessingError(f"Error extracting transactions: {str(e)}")


def export_to_excel(table: Table, output_document_format: OutputDocumentFormat) -> None:
    """
    Export table to Excel file.

    Args:
        table: Table to export
        output_document_format: Output format configuration

    Raises:
        ExportError: If export fails
    """
    try:
        if not table:
            raise ExportError("Cannot export empty table")

        # Check if we have the right number of columns
        if table and len(table[0]) != len(output_document_format.columns):
            raise ExportError(
                f"Table has {len(table[0])} columns but {len(output_document_format.columns)} "
                f"column names were provided"
            )

        df = pd.DataFrame(table, columns=output_document_format.columns)
        df.to_excel(output_document_format.path, index=False)  # type: ignore
    except ExportError:
        raise
    except Exception as e:
        raise ExportError(f"Failed to export to Excel: {str(e)}")


def replace_whitespace(text: Cell) -> Cell:
    """Replace whitespace in text."""
    if not isinstance(text, str):
        raise ValueError(f"Expected string, got {type(text).__name__}")
    return text.replace(" ", "")


def convert_to_float(text: Cell) -> Cell:
    """Convert text to float."""
    try:
        if not isinstance(text, str):
            raise ValueError(f"Expected string, got {type(text).__name__}")
        if not text:
            return 0.0
        text = text.replace(" ", "").replace(",", ".")
        return float(text)
    except ValueError:
        raise ValueError(f"Cannot convert '{text}' to float")


def setup_configuration(
    input_path: str, output_path: str
) -> Tuple[LoadingConfig, ProcessingConfiguration, ExportConfig]:
    """
    Set up the processing configuration.

    Args:
        input_path: Path to input DOCX file
        output_path: Path to output Excel file

    Returns:
        Tuple of (loading_config, processing_config, export_config)

    Raises:
        ValueError: If configuration setup fails
    """
    try:
        # Transaction row parsing configuration
        transaction_row_parsing_config = TransactionRowParsingConfig(
            field_count=3,
            id_test_func=lambda s: s.isdigit(),  # type: ignore
        )

        # Table format
        table_format = TableFormat(
            header_len=3,
            footer_len=2,
            account_cell_index=4,
            debit_cell_index=5,
            credit_cell_index=6,
            transaction_row_parsing_config=transaction_row_parsing_config,
        )

        # Input document format
        document_format = InputDocumentFormat(table_index=2, path=input_path)

        # Loading configuration
        loading_config = LoadingConfig(
            loading_strategy=load_xml_document,
            table_choose_strategy=choose_table,
            document_format=document_format,
        )

        # Processing configuration
        processing_config = ProcessingConfiguration(
            header_processing_strategy=empty_header,
            footer_processing_strategy=empty_footer,
            # detail_row_processing_strategy=ft.partial(
            #     process_detail_row_and_process_account,
            #     process_func=replace_whitespace,
            # ),
            detail_row_processing_strategy=ft.partial(
                process_detail_row_and_process_account_debit_credit,
                process_account_func=replace_whitespace,
                process_debit_func=convert_to_float,
                process_credit_func=convert_to_float,
            ),
            transaction_row_processing_strategy=parse_transaction_description,
            combine_rows_strategy=combine_rows,
            table_format=table_format,
        )

        # Output format
        output_document_format = OutputDocumentFormat(
            path=output_path,
            columns=[
                "Дата и время совершения текущей операции",
                "№ док.",
                "Код опер",
                "Код",
                "Счет",
                "Дебет",
                "Кредит",
                "Контрагент",
                "УНП",
                "Назначение",
            ],
        )

        # Export configuration
        export_config = ExportConfig(
            export_strategy=export_to_excel,
            output_document_format=output_document_format,
        )

        return loading_config, processing_config, export_config

    except Exception as e:
        raise ValueError(f"Error in configuration setup: {str(e)}")


def main(input_path: str = "vpsk.docx", output_path: str = "output.xlsx") -> int:
    """
    Main function to process document.

    Args:
        input_path: Path to input DOCX file
        output_path: Path to output Excel file

    Returns:
        Exit code (0 for success, non-zero for error)
    """
    try:
        # Setup configuration
        loading_config, processing_config, export_config = setup_configuration(
            input_path, output_path
        )

        # Load the document
        print(f"Loading document: {loading_config.document_format.path}")
        document = loading_config.loading_strategy(loading_config.document_format)

        # Choose the table
        print(f"Selecting table {loading_config.document_format.table_index}")
        table = loading_config.table_choose_strategy(
            document, loading_config.document_format
        )

        # Process the table
        print("Processing transactions...")
        transactions = extract_transactions(processing_config, table)

        # Export the result
        print(f"Exporting to {export_config.output_document_format.path}")
        export_config.export_strategy(
            transactions, export_config.output_document_format
        )

        print(
            f"✓ Successfully processed document and exported to {export_config.output_document_format.path}"
        )

        return 0

    except DocumentLoadingError as e:
        print(f"❌ Document loading error: {str(e)}")
        return 1
    except TableProcessingError as e:
        print(f"❌ Table processing error: {str(e)}")
        return 1
    except ExportError as e:
        print(f"❌ Export error: {str(e)}")
        return 1
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")
        return 1


if __name__ == "__main__":
    exit_code = main()
    exit(exit_code)
