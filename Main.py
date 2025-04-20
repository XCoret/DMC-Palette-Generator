import os  # Import OS for file and directory operations
import numpy as np  # Import NumPy for numerical operations
import pandas as pd  # Import pandas for data manipulation and analysis

from Utils.Constants import *  # Import constants 
from Utils.Image import readImage, getRoi  # Import readImage and getRoi functions
from Utils.Excel import createExcel  # Import createExcel function
from Utils.Color import getDMCColors, rgba2hex, nearestDMC  # Import color-related functions

if __name__ == '__main__':
    # Create results directory if it doesn't exist
    if not os.path.exists(RESULTS_FOLDER):
        os.makedirs(RESULTS_FOLDER) 
    
    # Get all PNG files from the sprites folder
    sprite_files = [os.path.join(SPRITES_FOLDER, f) 
                   for f in os.listdir(SPRITES_FOLDER) 
                   if f.endswith('.png')]
    
    # Load DMC color database
    colors_df = getDMCColors(DMC_FILE)
    
    # Process each sprite file
    for sprite_file in sprite_files:
        filename = os.path.basename(sprite_file).split('.')[0]
        print(f"Processing file {filename}...")
        
        # Load and crop the sprite image to its visible area
        sprite = getRoi(readImage(sprite_file))
        
        # Dictionary to track unique DMC colors (key: floss number, value: color info)
        unique_dmc = {}  
        temp_colors = []  # Temporary storage for color data
        index_counter = 0  # Counter for color indices (A, B, C...)
        
        # First pass: Identify all unique DMC colors in the sprite
        for i in range(sprite.shape[0]):
            for j in range(sprite.shape[1]):
                if sprite[i, j, 3] != 0:  # Check for non-transparent pixels
                    color = sprite[i, j]
                    dmc, floss = nearestDMC(color, colors_df)
                    
                    # If this DMC color hasn't been seen before
                    if floss not in unique_dmc:
                        # Assign next letter (A=65 in ASCII)
                        index = chr(65 + index_counter)  
                        unique_dmc[floss] = {
                            'INDEX': index,
                            'DMC': dmc,
                            'REAL': color
                        }
                        temp_colors.append([index, color, dmc, floss])
                        index_counter += 1
        
        # Create DataFrame from collected color data
        sprite_colors = pd.DataFrame(temp_colors, 
                                   columns=['INDEX', 'REAL', 'DMC', 'FLOSS'])
        
        # Initialize matrices for Excel output
        index_matrix = np.full((sprite.shape[0], sprite.shape[1]), "x", dtype=object)
        real_matrix = np.full((sprite.shape[0], sprite.shape[1]), "x", dtype="U9")
        dmc_matrix = np.full((sprite.shape[0], sprite.shape[1]), "x", dtype="U9")
        
        # Second pass: Fill matrices with color information
        for i in range(sprite.shape[0]):
            for j in range(sprite.shape[1]):
                if sprite[i, j, 3] != 0:  # Non-transparent pixels only
                    color = sprite[i, j]
                    dmc, floss = nearestDMC(color, colors_df)
                    index = unique_dmc[floss]['INDEX']
                    
                    # Populate matrices
                    index_matrix[i, j] = index
                    real_matrix[i, j] = rgba2hex(color)
                    dmc_matrix[i, j] = rgba2hex(dmc)
        
        # Generate Excel file with color information
        createExcel(index_matrix, real_matrix, dmc_matrix, sprite_colors, 
                   os.path.join(RESULTS_FOLDER, f"{filename}.xlsx"))