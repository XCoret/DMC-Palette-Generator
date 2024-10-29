import os  # Import OS for file and directory operations
import numpy as np  # Import NumPy for numerical operations
import pandas as pd  # Import pandas for data manipulation and analysis

from Utils.Constants import *  # Import constants 
from Utils.Image import readImage, getRoi  # Import readImage and getRoi functions
from Utils.Excel import createExcel  # Import createExcel function
from Utils.Color import getDMCColors, rgba2hex, nearestDMC, createMatrix # Import getDMCColors and rgba2hex functions

if __name__ == '__main__':
    # Check if the results folder exists, if not, create it
    
    if not os.path.exists(RESULTS_FOLDER):
        os.makedirs(RESULTS_FOLDER) 
    # Get a list of sprite files from the sprites folder
    sprite_files = [os.path.join(SPRITES_FOLDER, f) for f in os.listdir(SPRITES_FOLDER) if f.endswith('.png')]
    
    # Read DMC colors from CSV file and convert them to a suitable format
    colors_df = getDMCColors(DMC_FILE)
    
    # Process each sprite file
    for sprite_file in sprite_files:
        filename = os.path.basename(sprite_file).split('.')[0]  # Get the filename without extension
        print(f"Processing file {filename}...")  # Print the filename being processed
        # Read and process the sprite image
        sprite = getRoi(readImage(sprite_file))  # Get the region of interest (ROI) from the sprite
        unique_colors = []  # List to store unique colors found in the sprite
        temp_colors = []  # Temporary list to store color information
        
        # Iterate through each pixel of the sprite
        for i in range(sprite.shape[0]):
            for j in range(sprite.shape[1]):
                if sprite[i, j, 3] != 0:  # Check if pixel is not transparent
                    color = sprite[i, j]  # Get the RGBA color
                    # Check if the color is already in the unique colors list
                    if not any(np.array_equal(color, uc) for uc in unique_colors):
                        dmc, floss = nearestDMC(color, colors_df)  # Find nearest DMC color
                        # Append color details to temp_colors list
                        temp_colors.append([chr(len(unique_colors) + 65), color, dmc, floss])
                        unique_colors.append(color)  # Add color to unique list
        
        # Create DataFrame for sprite colors with color details
        sprite_colors = pd.DataFrame(temp_colors, columns=['INDEX', 'REAL', 'DMC', 'FLOSS'])

        # Create matrices for the sprite colors
        index_matrix, real_matrix, dmc_matrix = createMatrix(sprite, sprite_colors)
        # Create Excel file with the color matrices and details
        createExcel(index_matrix, real_matrix, dmc_matrix, sprite_colors, os.path.join(RESULTS_FOLDER, f"{filename}.xlsx"))
