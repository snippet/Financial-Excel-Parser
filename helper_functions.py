from datetime import datetime
from pathlib import Path
from typing import Any, List, Union

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

def serialize_value(cell: Cell) -> Union[str, float, int, None]:
    value = cell.value
    if value is None:
        return None
    return value

def remove_empty_rows(
    record: list[Union[str, float, int, None]]
) -> Union[list[Union[str, float, int, None]], None]:
    """
    Remove rows where all values are None.

    Args:
    record (list): A list of cell values.

    Returns:
    list or None: The original list if not all values are None, otherwise None.
    """
    if all(value is None for value in record):
        return None
    return record

def is_header_row(row: List[Cell]) -> bool:
    """
    Identify if a row is a header based on having more than one cell with datetime values.

    Args:
    row (List[Cell]): A list of cells in the row.

    Returns:
    bool: True if the row is identified as a header, otherwise False.
    """
    datetime_count = sum(1 for cell in row if isinstance(cell.value, datetime))
    return datetime_count > 1

def get_dataset_fields_idxs(row: List[Cell]) -> List[int]:
    """
    Get the indices of non-None cells in a row.

    Args:
    row (List[Cell]): A list of cells in the row.

    Returns:
    List[int]: A list of indices for cells that are not None.
    """
    indices = [idx for idx, cell in enumerate(row) if cell.value is not None]

    return list(range(indices[0], indices[len(indices) - 1] + 1))

def get_fields_type(
    row: List[Cell], current_types: List[Union[type, None]], field_indices: List[int]
) -> List[Union[type, None]]:
    """
    Update the current_dataset_fields_type with the types of the fields based on the indices.

    Args:
    row (List[Cell]): A list of cells in the row.
    current_types (List[Union[type, None]]): The current list of types.
    field_indices (List[int]): The indices of the fields to check.

    Returns:
    List[Union[type, None]]: Updated list of types.
    """
    if len(current_types) <= 0:
        current_types = [None] * (field_indices[-1] + 1)

    for idx in field_indices:
        cell_value = row[idx].value
        if cell_value is not None:
            current_type = type(cell_value)
            if current_types[idx] is None:
                current_types[idx] = current_type
    return current_types

def is_complex_header_row(
    prev_row_values: List[Cell],
    current_row: List[Cell],
    dataset_fields_idxs: List[Union[type, None]],
    dataset_fields_type: List[Union[type, None]],
) -> bool:
    """
    Identify if a row is a complex header based on the previous row being blank (all None) and type differences.

    Args:
    prev_row_values (List[Cell]): The previous row of cells.
    current_row (List[Cell]): The current row of cells.
    dataset_fields_type (List[Union[type, None]]): The list of types for the current data group fields.

    Returns:
    bool: True if the previous row is blank and the current row meets the type conditions.
    """
    # Check if prev row is blank
    if all(cell is None for cell in prev_row_values):
        # Check if all cells types of current_row are String or Datetime (ignore None fields)
        for cell in current_row:
            if cell.value is not None and not isinstance(cell.value, (str, datetime)):
                return False

        # Check if dataset_fields_type have different types than current_row (ignore None fields for both sides)
        for idx in dataset_fields_idxs:
            cell = current_row[idx]
            if cell is not None and cell.value is not None:
                current_type = type(cell.value)
                if (
                    dataset_fields_type[idx] is not None
                    and dataset_fields_type[idx] != current_type
                ):
                    # print(dataset_fields_type[idx], current_type)`
                    return True

    return False

def set_end_row_to_prev_dataset(fields_to_check, end_row_value, dataset):
    # Add end_row to the previous dataset if there is at least one matching field
    for entry in dataset:
        if "end_row" not in entry:
            if any(field in entry["fields"] for field in fields_to_check):
                entry["end_row"] = end_row_value
                return dataset
    return dataset

def process_table(
    ws: Worksheet,
) -> List[dict[str, Union[str, float, int, None]]]:
    """
    process_simple_table handles a simple spreadsheet which has one table starting from the top left corner.
    Its first row is its header and the following rows are data records.
    """
    records = []
    datasets_range = []
    current_dataset_fields_idxs = []
    current_dataset_fields_type = []

    prev_row_values = [None]

    rows = list(ws.iter_rows())
    for idx, row in enumerate(rows):
        values = [serialize_value(cell) for cell in row]
        records.append(values)

        # Check if the row is Header
        if is_header_row(row) | is_complex_header_row(prev_row_values,row,current_dataset_fields_idxs,current_dataset_fields_type):
            current_dataset_fields_type = [] # reset types

            # Check if the row has multi datasets range
            prev_dataset_fields_idxs = current_dataset_fields_idxs
            current_dataset_fields_idxs = get_dataset_fields_idxs(row)

            is_multi_dataset = ((abs(len(prev_dataset_fields_idxs) - len(get_dataset_fields_idxs(row))) > 3) & (len(prev_dataset_fields_idxs) != 0))

            if(is_multi_dataset):
                # ignore fields from other dataset in the same row.
                current_dataset_fields_idxs = [
                    item
                    for item in current_dataset_fields_idxs
                    if item not in prev_dataset_fields_idxs
                ]

            # Ajust the position of first field in dataset (looking for a labels column)
            for x in range((current_dataset_fields_idxs[0] - 2), (current_dataset_fields_idxs[0] + 3)):
                first_field_idx = x
                string_fields = 0

                for i in range(1, 7):
                    if idx + i < len(rows):
                        next_row = rows[idx + i]
                        if len(next_row) > first_field_idx:
                            string_fields += 1 if isinstance(next_row[first_field_idx].value, str) else 0

                if(string_fields >= 3):
                    # when find the idx of labels column, set the idx as a fisrt field idx in dataset..ignoring messy columns
                    # print(f"{x} {(current_dataset_fields_idxs[0] - 2)}, {(current_dataset_fields_idxs[0] + 3)} ({current_dataset_fields_idxs})")

                    if x in current_dataset_fields_idxs:
                        current_dataset_fields_idxs = current_dataset_fields_idxs[current_dataset_fields_idxs.index(x):]
                    else:
                        current_dataset_fields_idxs.insert(0, x)

            # Set the end row of dataset
            datasets_range = set_end_row_to_prev_dataset(current_dataset_fields_idxs, idx-1, datasets_range)

            # Append dataset range
            datasets_range.append(
                {
                    "row": idx,
                    "fields": current_dataset_fields_idxs,
                }
            )

        # body of dataset
        else:
            if len(current_dataset_fields_idxs) > 1:
                current_dataset_fields_type = get_fields_type(row, current_dataset_fields_type, current_dataset_fields_idxs)

        prev_row_values = values

    return (
        records,
        datasets_range,
    )

def calculate_num_leading_space(current_header: str) -> int:
    return len(current_header) - len(current_header.lstrip())

def get_hierarchical_string(hierarchical_levels, cell_value):
    level = calculate_num_leading_space(cell_value)
    keys = sorted(hierarchical_levels.keys())
    result = []

    for key in keys:
        if key < level:
            result.append(hierarchical_levels[key].strip())

    result.append(cell_value.strip())

    return " > ".join(result)
