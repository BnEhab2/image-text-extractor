# Image Text Extractor using Google Lens

This project automates the process of extracting text from images using **Google Lens** and **Selenium**, then saves the extracted text into a **Word document** with **Arabic RTL support**.

## 📦 Features

- Uploads images to [Google Lens](https://lens.google.com/upload )
- Extracts text using Selenium and clipboard (`pyperclip`)
- Saves extracted text in a `.docx` file
- Supports **Arabic RTL** formatting

## 🧰 Requirements

- Python 3.8+
- Microsoft Edge (or Chrome)
- `msedgedriver.exe` (or chromedriver)
- Images folder with `.jpg`, `.png`, or `.jpeg` files
- Folder that has images
  
## 📥 Installation

```bash
pip install selenium pyperclip python-docx
