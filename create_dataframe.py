"""Module for creating sample pandas DataFrames."""

import pandas as pd
import numpy as np


def create_sample_dataframe(rows: int = 10, cols: int = 5) -> pd.DataFrame:
    """Create a sample DataFrame with random numeric data.

    Args:
        rows: Number of rows. Defaults to 10.
        cols: Number of columns. Defaults to 5.

    Returns:
        A DataFrame with random integer values.
    """
    column_names = [f"col_{i + 1}" for i in range(cols)]
    data = np.random.randint(0, 100, size=(rows, cols))
    return pd.DataFrame(data, columns=column_names)


if __name__ == "__main__":
    df = create_sample_dataframe()
    print(df)
