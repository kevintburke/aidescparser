# AI Serial Item Description Parser
This program uses a locally hosted AI model to parse unstructured item descriptions for issues of serials into structured enumeration/chronology data to be loaded into Alma. The script in its current form uses Ollama to run Qwen2.5-coder to handle the data and outputs an xlsx file valid for loading using the Alma <a href="https://developers.exlibrisgroup.com/appcenter/item-updater-by-excel/" title="Item Updater by Excel" target="_blank">Item Updater by Excel Cloud App</a>.

## Inputs
- Spreadsheet (xlsx or csv) with item descriptions to be parsed and item identifiers
  - Identifiers can be either the item barcode or a combination of MMS ID, holdings ID, and item PID

## Outputs
- Spreadsheet (xlsx) with validated parsed descriptions for loading into Alma
- Spreadsheet (xlsx) with item IDs and descriptions for items that failed parsing
- Spreadsheet (csv) with logs for each AI call including length and pass/fail

## Process
1. Input spreadsheet is loaded and validated
2. User is prompted to select a method for identifying items (either barcode or combination MMS/holdings ID/item ID)
3. User is prompted to select column(s) containing item identifier(s) and column containing item descriptions
4. Script loads the base prompt and checks that Ollama is active and Qwen2.5-coder is available
5. Each line of the input is fed to the AI model alongside the complete prompt
6. AI response is validated for structure (Python dictionary with the keys Enum A-D and Chron I-K)
7. AI response is checked against common issues (dates in Enum A, fields filled out not according to hierarchy, missing/added numeric data in parsed description, etc.) using a set of regular expressions
8. User is prompted to save the successful outputs and the failed outputs to two separate files with columns based on the item ID method chosen in step 2
9. Log is saved to the working directory with a default name including a timestamp for easy review

## Notes
- The prompt must be stored in the same directory as the script, which must be the current working directory when running the script
- The complete prompt is fed to the AI model at each call because I wasn't able to find a way to maintain session memory between calls when running Ollama through Python. This causes a significant loss in performance, so should be addressed in future versions when possible.
- Currently the outputs of the AI model are validated using a set of regular expressions to target common errors identified in testing. Testing was done using additional calls to Qwen to validate its own responses, but the results were found not to be worth the added runtime. A more specialized classifier-type model, such as Laya, could possibly improve performance of validation.
- Current testing shows the rate of false positives to be around 2% with false negatives around 1%. The overall success rate is around 90%. These figures vary widely with the consistency of the input data, though, so should be treated as rough estimates only.
