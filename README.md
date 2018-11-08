# HL7_SIM
Generator of HL7 simulation files

# create executable using pyinstaller
pyinstaller --windowed --icon=tool.ico PurplePanda.py

# to add png resource files to be used in tool
# first add the file to directory
# add the file in img,qrc
# create py file using pyrcc4
pyrcc4 img.qrc -o img_qr.py
# now the resource can be used by importing img_qr.py file