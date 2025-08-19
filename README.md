# PDF to PowerPoint Converter

A simple desktop tool to convert PDF files into PowerPoint presentations (PPTX), built with Python and Tkinter.

---

## Features

- **High-quality image conversion**: Each PDF page is rendered as an image and placed onto a PowerPoint slide.
- **Maintains aspect ratio**: Slides are sized to match the PDF page dimensions for accurate rendering.
- **Customizable DPI**: Adjust the image quality of each slide.
- **Adjustable slide dimensions**: Set the slide width (in inches) as desired.
- **Responsive UI**: Conversion occurs in a background thread, keeping the interface responsive.
- **Conversion log & status**: Real-time feedback and logging during conversion.
- **Progress bar and cancellation**: Visual progress and option to cancel ongoing conversion.
- **About dialog**: Built-in help and information dialog.

---

## Requirements

- **Python**: 3.8 or higher
- **Dependencies**:
  - `PyMuPDF` (`fitz`)
  - `python-pptx`
  - `tkinter` (bundled with most Python distributions)
- Install dependencies with:
  ```bash
  pip install PyMuPDF python-pptx
  ```

---

## Usage

1. **Run the application**:
   ```bash
   python pdf_to_ppt.py
   ```
   *(Replace `pdf_to_ppt.py` with your filename if different.)*

2. **Select a PDF**:
   - Click "Browse" next to "Select PDF File" to choose your input PDF.

3. **Set output location**:
   - Click "Browse" next to "Output PPT File" to choose where to save the PowerPoint file.

4. **Adjust settings**:
   - Set the DPI (image quality) and slide width as desired.

5. **Convert**:
   - Click "Convert PDF to PPT" and watch the progress.
   - You can cancel conversion at any time.

6. **Logs and status**:
   - View real-time logs and status messages in the main window.

7. **About**:
   - Use the Help > About menu for app information and requirements.

---

## Notes

- **Temporary PNGs**: The converter saves each page as a temporary PNG image before inserting it into PowerPoint. These files are cleaned up automatically.
- **Aspect Ratio**: The slide height is adjusted automatically to maintain the PDF's original aspect ratio.
- **Threading**: Conversion is performed in a background thread to keep the UI responsive.

---

## Troubleshooting

- **Missing Libraries**: If you get an error about missing libraries (`fitz`, `pptx`), make sure to install them using pip.
- **Python Version Error**: The app requires Python 3.8+. If you see a version error, upgrade your Python installation.
- **Permission Issues**: Ensure you have write permissions to the selected output directory.

---

## License

This tool is provided as-is, without warranty.

---

## Credits

- Built using:
  - [PyMuPDF](https://github.com/pymupdf/PyMuPDF)
  - [python-pptx](https://github.com/scanny/python-pptx)
  - [Tkinter](https://docs.python.org/3/library/tkinter.html)

---

## Screenshot
<img width="609" height="526" alt="Screenshot 2025-08-19 at 6 09 56 PM" src="https://github.com/user-attachments/assets/12823a96-8b52-4297-b41f-0c87e5cfb7d5" />



📜 License This project is licensed under the MIT License – see the LICENSE file for details.
