# Motor torque map Excel import

In **Engineering Report**, select an Electric or Hybrid configuration and open
**Motor Performance Curves (Optional)**. Enable **Include one shared motor
capability curve**, then use **Upload Motor Torque Map (.xlsx)**. This motor
workbook is separate from the VSM simulation results uploaded at the top of the
application. Clipboard operations are no longer required for motor-map import.

Download **Download blank motor torque map template (.xlsx)** to prepare a map.
The template contains only headers and unit comments, with no example ratings.
Use your motor manufacturer's actual data:

| Column in row 1 | Required values | Units |
| --- | --- | --- |
| Motor RPM | Finite, nonnegative, unique motor shaft speed | rpm |
| Peak Torque | Finite, nonnegative peak torque magnitude | Nm |
| Cont Torque | Finite, strictly positive continuous torque magnitude | Nm |

Enter at least two complete data rows immediately below the headers. Column
order can vary. Case, whitespace and minor header variations are accepted,
including `Motor speed [rpm]`, `Motor peak torque [Nm]` and
`Motor continuous torque [Nm]`. Bracketed or parenthesized units must be rpm
for speed and Nm for torque; N.m and N·m are also accepted. Headers without
units have the units stated above. No unit conversion or unit inference from
the values is performed. Convert other units before upload.

Save as an unencrypted `.xlsx` workbook containing numeric values. Older `.xls`,
CSV and renamed non-Excel files are not supported for this motor import. Formula
cells are rejected; export their numeric values into the motor workbook first.
Microsoft Excel does not need to be installed on the machine running the tool.

The importer finds worksheets whose first row contains the three required
headers, ignoring unrelated sheets. If several match, choose **Motor torque
worksheet**. If none match, select the intended worksheet when offered and
correct the displayed header errors. Unrelated extra columns are allowed;
do not put titles, unit-only rows, totals or footnotes in the data rows.

Check the import confirmation, RPM-point count, preview table and torque curves.
RPM is sorted in increasing order with both torque values kept paired. Entirely
blank rows, including trailing rows, are ignored. Missing values, nonnumeric
values, duplicates, invalid signs and ambiguous columns produce an actionable
error, with a worksheet/cell location for invalid data. No incomplete engineering
row is discarded or replaced with zero. Fix the workbook, save it and upload it
again. Reports are blocked while the enabled upload is invalid; a previous
successful workbook is not used instead.

Enter the known **Final-drive ratio** and **Driveline efficiency** to include peak
and wheel-side capability overlays and the combined Motor Torque and Power
chart/slide. Continuous motor-side assessment can operate without these two
parameters. The importer adds no new peak-versus-continuous relationship rule;
the existing independent curve semantics remain in effect.

Proceed with **STEP 3 - Validate** and **STEP 4 - Generate Engineering Report**,
then download the existing Excel and PowerPoint reports. The supplied reference
curve is shared by the four motors, as before. Existing generator symmetry,
linear interpolation, operating limits and no-extrapolation policies are unchanged.

Uploads stay attached to the current VSM source content/path and reporting
profile during ordinary widget reruns. Uploading a replacement reparses its
contents, even if its filename is unchanged; worksheet selection resets when
the contents change. Selecting another source/profile starts without inherited
specifications. **Reuse previous motor specifications for this source** is an
explicit transfer, and **Clear motor curve** clears the current inputs. Session
inputs are not saved across application restarts; keep your workbook.

If no workbook is uploaded, the existing manual table entry/paste fallback is
still available, including independent peak-only or continuous-only data. Remove
an uploaded workbook to use that fallback. Diesel does not use this input section.
