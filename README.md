# JSON to CSV/Excel Converter

This script converts JSON data into CSV or Excel format. You can provide the source JSON file and specify the desired output format to generate the required output.
The output is placed in target directory based on the format selected target_csv/target_excel
## Usage

1. **Environment Variables**

   Before using the script, set the following environment variables in your `.env` file:

   ```env
   INPUT_FILE=filename.json
   TARGET_FORMAT=format_to_conver (csv/excel) supported

- INPUT_FILE: Filename of the  source JSON document. File should be in the root directory
- TARGET_FORMAT: Desired output format (either "csv" or "excel").

## Source JSON Specifications
The script expects the source JSON to be provided through a file. The JSON structure should match the following format:


```json
[
   {
      "ParameterInformation": [
         {
            "ParameterName": "Wirkleistung",
            ...
         }
      ],
      "ErrorInformation": [],
      "FamilyLoaded": true,
      "FileName": "Family1.rfa"
   }
]
```




- Each element in the JSON array represents a specific file.
- The "ParameterInformation" array contains parameter details. For Ex:

```json
[
      {
        "ParameterName": "Wirkleistung",
        "ParameterValues": {
          "Allgemein": "0"
        },
        "Unit": "BSU_WATTS",
        "ParameterType": "Double",
        "ParameterValueType": "ElectricalPower",
        "IsSharedParameter": true,
        "IsInstanceParameter": true,
        "IsReportingParameter": false,
        "ParameterGroup": "BS_ELECTRICAL_CIRCUITING",
        "Formula": null,
        "IsDeterminedByFormula": false,
        "IsEnumType": false,
        "EnumValues": [],
        "Guid": "da6a3bbc-a600-40fa-970a-08fb75f19111",
        "Description": null,
        "BuiltInParameterEnumName": "INVALID"
      },
]
```

## Example
Suppose you have a JSON file named input.json and want to convert it to a CSV file. Your .env file should contain:

````
INPUT_FILE=input.json
TARGET_FORMAT=csv
````


## Dependencies
Python 3.x
Required Python packages can be installed using `pip install -r requirements.txt`.

