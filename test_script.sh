#!/bin/bash

# Define test file names
EXCEL1="test_Excel1.xlsx"
UPDATED_EXCEL2="test_updated_excel2.xlsx"
PYTHON_SCRIPT="process_data.py"

# Remove existing test files if they exist
echo "Cleaning up old test files..."
rm -f "$EXCEL1" "$UPDATED_EXCEL2"

# Create test Excel1.xlsx using Python
echo "Creating test Excel1.xlsx..."
python3 - <<EOF
import pandas as pd

data1 = {
    "Well Id": ["W001", "W002", "W003"],
    "Date of Sampling": ["2024-03-25", "2024-03-26", "2024-03-27"],
    "Area": ["North", "South", "East"],
    "Parameter": ["pH", "Turbidity", "Chlorine"],
    "Unit": ["pH Units", "NTU", "mg/L"],
    "Result": [7.2, 3.5, 1.1]
}
df1 = pd.DataFrame(data1)
df1.to_excel("$EXCEL1", index=False)
EOF

echo "Created test_Excel1.xlsx"

# Create test updated_excel2.xlsx using Python
echo "Creating test updated_excel2.xlsx..."
python3 - <<EOF
import pandas as pd

data2 = {
    "Name": ["W001", "W002"],  # Same as Excel1 for testing duplicates
    "TimeStamp": ["2024-03-25T12:00:00", "2024-03-26T12:00:00"],
    "SourceTimeStamp": ["2024-03-25T12:00:00", "2024-03-26T12:00:00"],
    "Area": ["North", "South"],
    "Latitude": [None, None],  # Empty fields in test data
    "Longitude": [None, None],
    "Type": ["pH", "Turbidity"],
    "Unit": ["pH Units", "NTU"],
    "Value": [7.2, 3.5],
    "Relationships": [None, None]
}
df2 = pd.DataFrame(data2)
df2.to_excel("$UPDATED_EXCEL2", index=False)
EOF

echo "Created test_updated_excel2.xlsx"

# Run the main Python script
echo "Running process_data.py..."
python3 "$PYTHON_SCRIPT"

# Verify the updated file
echo "Verifying results..."
python3 - <<EOF
import pandas as pd

try:
    df = pd.read_excel("$UPDATED_EXCEL2")
    print("\n### Final Data in updated_excel2.xlsx ###")
    print(df)
except Exception as e:
    print("Error reading updated file:", e)
EOF

echo "Test completed!"
