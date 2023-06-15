import os
import subprocess
import argparse
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk


def pdf_to_images(pdf_path):
    """
    Function to convert PDF to images
    """
    return convert_from_path(pdf_path)


def create_docx_from_images(images, docx_path, width, progress=None, root=None):
    """
    Function to create Word document from images
    """
    doc = Document()

    for i in tqdm(range(len(images)), desc='Converting images to Word document', unit='image'):
        img_filename = f'image_{i}.png'
        images[i].save(img_filename)  # saving images
        doc.add_picture(img_filename, width=Inches(width))  # Add image to the docx and adjust the size as per your requirement
        os.remove(img_filename)  # Remove image file after using
        if progress:
            progress['value'] += 100/len(images)
            if root:
                root.update_idletasks()

    doc.save(docx_path)


def pdf_to_docx(pdf_path, docx_path, width, progress=None, root=None):
    """
    Function to convert PDF to Word document
    """
    images = pdf_to_images(pdf_path)
    create_docx_from_images(images, docx_path, width, progress, root)


def open_file(path):
    """
    Function to open a file depending on the OS
    """
    if os.name == 'nt':  # for windows
        os.startfile(path)
    else:
        opener = "open" if os.name == "posix" else "xdg-open"  # for mac and linux respectively
        subprocess.call([opener, path])


def main():
    """
    Main function to handle arguments
    """
    parser = argparse.ArgumentParser(description='Convert a PDF to a Word document with each page as an image.')
    parser.add_argument('pdf_path', nargs='?', default=None, help='The path to the input PDF file')
    parser.add_argument('docx_path', nargs='?', default=None, help='The path to output the Word document')
    parser.add_argument('-w', '--width', type=float, default=6.0,
                        help='The width of the images in the Word document (in inches, default: 6.0)')

    args = parser.parse_args()

    if args.pdf_path and args.docx_path:
        pdf_to_docx(args.pdf_path, args.docx_path, args.width)
        open_file(args.docx_path)
    else:
        root = tk.Tk()
        root.withdraw()  # hide the main window

        pdf_path = filedialog.askopenfilename(title="Select PDF file",
                                              filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        docx_path = filedialog.asksaveasfilename(title="Save Word file",
                                                 filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
                                                 defaultextension=".docx")

        width = simpledialog.askfloat("Input", "Enter the image width in inches:", minvalue=0.0, initialvalue=6.0)

        root.title('PDF to DOCX Conversion')
        progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        progress.pack()

        root.deiconify()  # show the main window

        pdf_to_docx(pdf_path, docx_path, width, progress, root)
        open_file(docx_path)

        root.destroy()  # close the main window


if __name__ == "__main__":
    main()
