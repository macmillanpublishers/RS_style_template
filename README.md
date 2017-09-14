# RS_style_template
Revised version of Macmillan styles for Word manuscripts, to be deployed with RSuite.

## Editing Styles
All data about the styles is stored in *WordTemplateStyles.xlsm*. This includes validation of the entries. See rows 1 and 2 for more detail.

## Generating Style Template
The actual style objects are stored in *macmillan.dotx*. To generate a new *macmillan.dotx*, run *CreateMacmillanTemplate.bat* (PC only), or
open *StyleTemplateCreator.docm* in Word and run macros *WriteTemplatefromJson* and *WriteNoColorTemplatefromJson"*

This also writes all of the data in WordTemplateStyles.xlsm to *macmillan.json* so we can track it in a non-binary format.