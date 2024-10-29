# Sprite Color Mapping to DMC

## Overview

This project processes sprite images to extract colors and map them to DMC (Dollfus-Mieg & Company) embroidery floss colors. It reads sprite images, identifies unique colors, finds the nearest DMC colors, and outputs the results into an Excel file.

## Features

- Extracts unique colors from sprite images.
- Maps each color to the nearest DMC color.
- Generates an Excel file with detailed color information.

## Requirements

- Python 3.x
- Libraries:
  - OpenCV
  - NumPy
  - pandas
  - openpyxl

## Installation

1. Clone the repository or download the code files.
2. Navigate to the project directory in your terminal.
3. Install the required libraries using pip:

   ```sh
   pip install -r requirements.txt
   ```
4. Ensure that you have a folder named `Sprites` containing your sprite images and a `DMC.csv` file with DMC colors.

## Usage

1. Place your sprite images (in PNG format) inside the `Sprites` folder.
2. Prepare a `DMC.csv` file with columns `R`, `G`, `B`, dmc, and Floss.
3. Run the script:
```sh
python Main.py
```
4. The results will be saved in the `Results` folder as Excel files named after each sprite.

## Functions

`nearestDMC`

Finds the nearest DMC color for a given color by calculating the Euclidean distance.

`rgba2hex`

Converts an RGBA color to a hexadecimal string.

`createMatrix`

Creates matrices for the sprite colors, including indices, real colors, and DMC colors.

`font_color`

Determines the appropriate font color based on the luminance of the input hexadecimal color.

`createExcel`

Generates an Excel file containing the color matrices and associated data, including formatting for better readability.

`getRoi`

Extracts the region of interest (ROI) from the image based on non-transparent pixels.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

 - OpenCV for image processing.
 - NumPy for numerical operations.
 - pandas for data manipulation.
 - openpyxl for creating Excel files.