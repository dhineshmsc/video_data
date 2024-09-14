import os
from pymediainfo import MediaInfo
import pandas as pd
from datetime import datetime

# Define the supported video formats
SUPPORTED_FORMATS = ('.mp4', '.avi', '.mkv', '.mov')

def get_video_properties(file_path):
    # Get file size in MB
    file_size = os.path.getsize(file_path) / (1024 * 1024)

    # Get file type (extension)
    file_type = os.path.splitext(file_path)[1]

    # Get video properties using pymediainfo
    media_info = MediaInfo.parse(file_path)
    duration = None
    width = None
    height = None

    for track in media_info.tracks:
        if track.track_type == 'Video':
            # Duration in seconds
            if track.duration:
                duration = float(track.duration) / 1000  # Convert from ms to seconds
            
            # Get frame width and height (resolution)
            width = track.width
            height = track.height

    # Format frame size as "Height x Width" or "Unknown" if not available
    frame_size = f"{height} x {width}" if width and height else "Unknown"

    return {
        "File Size (MB)": round(file_size, 2),
        "File Type": file_type,
        "Duration (seconds)": round(duration, 2) if duration else "Unknown",
        "Frame Size": frame_size  # Format "Height x Width"
    }

def process_video_files_in_folder(folder_path):
    if not os.path.exists(folder_path):
        return "Folder does not exist."

    # List to store properties of all video files
    video_files_properties = []

    # Iterate through all files in the folder
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)

            # Check if the file is a video by its extension
            if file_name.lower().endswith(SUPPORTED_FORMATS):
                properties = get_video_properties(file_path)
                video_files_properties.append({
                    "File Name": file_name,
                    "File Type": properties["File Type"],
                    "Frame Size": properties["Frame Size"],
                    "Duration (seconds)": properties["Duration (seconds)"],
                    "File Size (MB)": properties["File Size (MB)"]
                })

    return video_files_properties

def export_to_excel(df, base_filename):
    # Generate timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    # Create full filename with timestamp
    output_filename = f"{base_filename}_{timestamp}.xlsx"
    # Export the DataFrame to Excel
    df.to_excel(output_filename, index=False)
    return output_filename

def find_duplicates(video_properties):
    # Create a DataFrame from the video properties
    df = pd.DataFrame(video_properties)
    
    # Find duplicates based on File Size (MB) and Duration (seconds)
    duplicates = df[df.duplicated(subset=["File Size (MB)", "Duration (seconds)"], keep=False)]
    
    return duplicates

# Example usage
folder_path = 'D:\\Dhinesh\\Movies'  # Replace with your folder path
video_properties_base_filename = 'video_properties'  # Base filename for all video properties
duplicates_base_filename = 'duplicate'  # Base filename for duplicate videos

# Process video files and get their properties
video_properties = process_video_files_in_folder(folder_path)

# Export the video properties to an Excel file with timestamp
video_properties_path = export_to_excel(pd.DataFrame(video_properties), video_properties_base_filename)

# Find duplicates
duplicates = find_duplicates(video_properties)

# Export duplicates to an Excel file with timestamp if any are found, otherwise print a message
if not duplicates.empty:
    duplicates_path = export_to_excel(duplicates, duplicates_base_filename)
    print(f"Exported duplicate video properties to {duplicates_path}")
else:
    print("No duplicate files found.")

print(f"Exported all video properties to {video_properties_path}")
