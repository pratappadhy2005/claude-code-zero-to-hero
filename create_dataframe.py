import pandas as pd
import numpy as np

df = pd.DataFrame(
    np.random.randint(0, 100, size=(10, 5)),
    columns=["A", "B", "C", "D", "E"]
)

print(df)
