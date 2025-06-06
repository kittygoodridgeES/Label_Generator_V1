# Label_Generator_V1
Generate QR code labels for factory locations

Run main.py to generate gui

Setup:
- make sure printer drivers are downloaded
- edit line 319 in main.py to printer name saved on PC

To Note:
- This program is only able to generate labels on 50mm wide label paper

Generating Standard Labels:
- run main.py to open gui
- enter label requirements
- press generate labels - this will save a pdf called labels.pdf to the folder
- press print - to send labels to brother printer where they will be automatically printed and cut

Generating labels from csv:
- Generate csv with single column of IDs of labels to print
- upload csv to folder with gui inside
- make sure all other csv's deleted
- when run gui select magazine and ignore all other fields
