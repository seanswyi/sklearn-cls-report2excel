# sklearn-cls-report2excel
Repository containing script that converts the classification report from [`sklearn.metrics.classification_report`](https://github.com/scikit-learn/scikit-learn/blob/main/sklearn/metrics/_classification.py#L2405) to a formatted Excel spreadsheet.

## Usage

You can either use the function `convert_report2excel` inside a separate script so you're saving your formatted files as you go, or you can run `convert_report2excel.py` as a script itself and provide the directory containing the report files as an argument.

### Using `convert_report2excel` as a Function

Some things to keep in mind:
  1. Set `output_dict=True` in `classification_report` and feed the dictionary to the function.
  2. Instantiate a `openpyxl.Workbook` object first.
  3. If you don't want the default sheet, delete it.

```Python
import numpy as np
from openpyxl import Workbook
import pandas as pd
from sklearn.metrics import classification_report

from convert_report2excel import convert_report2excel


workbook = Workbook()
workbook.remove(workbook.active) # Delete default sheet.

y_true = np.array(['cat', 'dog', 'pig', 'cat', 'dog', 'pig'])
y_pred = np.array(['cat', 'pig', 'dog', 'cat', 'cat', 'dog'])

report = classification_report(
    y_true,
    y_pred,
    digits=4,
    zero_division=0,
    output_dict=True
)

workbook = convert_report2excel(
    workbook=workbook,
    report=report,
    sheet_name="animal_report"
)
workbook.save("animal_report.xlsx")
```

### Using `convert_report2excel` as a Script

If you don't specify the `--save_dir` argument, the results will be saved automatically to `--report_dir`.

```
python convert_report2excel.py --report_dir $PATH_TO_REPORTS
```
