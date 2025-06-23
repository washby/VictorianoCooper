# Image Color Extractor Application

This Python application provides a graphical user interface (GUI) to process JPG images within a selected folder. It crops a central rectangular region from each image, calculates various RGB color statistics for that region, and outputs the results to an Excel file. Additionally, it saves the cropped images to a new subfolder.

## Features

- Select a folder containing JPG images.
- Automatically crops a central rectangular region (default 15% of image dimensions).
- Calculates and records the following RGB statistics for the cropped region:
  - Mean (Average)
  - Median
  - Minimum
  - Maximum
  - First Quartile (Q1)
  - Third Quartile (Q3)
  - Mean of values within the Interquartile Range (IQR)
- Outputs a summary Excel file (image_colors.xlsx) with these statistics for each image.
- Saves the cropped images to a cropped_images subfolder, appending _cropped to the original filename.

# Installation

1. **Download Project**
    
    If you are not familar with Git, you can download the project by clicking the green `Code` button (above the list of files but below the main menu bar). When you download it, you will need to unzip it. Navigate to your Downloads directory and right click the file and select `Extract All.`

    Make sure you pay attention where you decide to extract the files.


2. **Install Python**

   If you don't already have Python installed, download the latest version from the official Python website:
[Download Python](https://www.python.org/downloads/)

    Follow the installation instructions for your operating system. ***Ensure you check the "Add Python to PATH" option during installation on Windows.***



3. **Install Project Dependencies**

   Once Python is installed, open your terminal or command prompt and navigate to the directory where you saved the image_color_extractor_app.py script. The easiest way to do this is to open File Explorer to where you downloaded and unzipped the project and right click the white area below the files and select "Open in Terminal." 

   Then, run the following command to install the required libraries:

        pip install -r requirements.txt

# How to Run the Program
Double click the image_color_extractor_app.py

A graphical user interface window will appear, allowing you to select your image folder and process the files.