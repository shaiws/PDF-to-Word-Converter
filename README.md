# PDF to Word Converter

This script allows users to convert a PDF file into a Word document. Each page of the PDF is converted into an image and then added to the Word document in the order they appear in the PDF. 

## Features

- Convert a PDF file into a Word document with each page as an image.
- Define the width of the images in the Word document.
- Progress bar to track the conversion process.
- Can handle arguments for the paths and width via the command line.
- If arguments are not supplied, a GUI will prompt you for the file paths and image width.

## Installation

Before running the script, you will need to install the necessary Python libraries. You can do this by running:

```bash
pip install -r requirements.txt
```


## Usage
Follow this before running: https://pdf2image.readthedocs.io/en/latest/installation.html

To use the script from the command line, use the following format:

```bash
python script.py input.pdf output.docx --width 5.0
```

- `input.pdf` is the path to the PDF file you want to convert.
- `output.docx` is the path where you want to save the Word document.
- `--width` is an optional argument where you can specify the width of the images in inches. If not provided, it defaults to 6.0.

If you run the script without any arguments:

```bash
python script.py
```

A GUI will open and prompt you to select the PDF file, the location to save the Word document, and the width of the images.

## Contributing

If you would like to contribute to this project, feel free to fork the repository and submit a pull request.

## License

This project is licensed under the MIT License.
