# Image Color Extractor Application
This Python application provides a graphical user interface (GUI) to process JPG images within a selected folder. It crops a central rectangular region from each image, calculates various RGB color statistics for that region, and outputs the results to an Excel file. Additionally, it saves the cropped images to a new subfolder.

Features
Select a folder containing JPG images.

Automatically crops a central rectangular region (default 15% of image dimensions).

Calculates and records the following RGB statistics for the cropped region:

Mean (Average)

Median

Minimum

Maximum

First Quartile (Q1)

Third Quartile (Q3)

Mean of values within the Interquartile Range (IQR)

Outputs a summary Excel file (image_colors.xlsx) with these statistics for each image.

Saves the cropped images to a cropped_images subfolder, appending _cropped to the original filename.

Installation
1. Install Python
   If you don't already have Python installed, download the latest version from the official Python website:

Download Python

Follow the installation instructions for your operating system. Ensure you check the "Add Python to PATH" option during installation on Windows.

2. Prepare requirements.txt
   The application relies on the following Python libraries:

Pillow (for image processing)

openpyxl (for Excel file manipulation)

numpy (for statistical calculations)

You should create a file named requirements.txt in the same directory as the Python script with the following content:

Pillow
openpyxl
numpy

3. Install Dependencies
   Once Python is installed, open your terminal or command prompt and navigate to the directory where you saved the image_color_extractor_app.py script and requirements.txt. Then, run the following command to install the required libraries:

pip install -r requirements.txt

How to Run the Program
Save the provided Python script (e.g., as image_color_extractor_app.py) in a directory on your computer.

Ensure you have created the requirements.txt file in the same directory.

Open your terminal or command prompt.

Navigate to the directory where you saved the script:

cd path/to/your/script/directory

(Replace path/to/your/script/directory with the actual path.)

Run the application using the Python interpreter:

python image_color_extractor_app.py

A graphical user interface window will appear, allowing you to select your image folder and process the files.