
# To build for MacOS
pyinstaller --windowed --onefile --name "WP Converter" --icon=icon.icns wpd_to_docx.py 

## Explanation

* --windowed prevents a Terminal window; 
* --onefile bundles everything into a single executable inside the .app; 
* --icon is optional (use a 1024Ã—1024 .icns).
