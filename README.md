# DataGovernance
> This is a tool developed in Python to assist with the AZ MDM data governance process, particularly during the migration project Mainframe>MDM>PIC. The team checks the integrity of the data and evaluate business rules are being fullfiled by synchronizing the data between the MDM platform and the current item information on Mainframe. This tool's purpose is to shorten the analysts' repetitive tasks of preparing a xlsx file to work on and running the respective queries.
This process has been subdivided into steps or "stages" in order to simplify the analysis done by the product integrity team. The stages are roughly as follows:
1. Evaluate the item's Global Attributes, mainly the POV ID and DCs along with item flags and key id data
2. Check Packaging data is alligned and complete according to the current item's package level (1,2 or 3)
3. Analyze and complete Pricing and Hazmat data for all valid items

In order for the tool to work, the user must export from the MDM platform a file including items with the particular scope for the stage the user wishes to work on.  
**At the moment only the first stage is available. Further releases will include the next stages.**
