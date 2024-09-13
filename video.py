import os
import pandas as pd
from moviepy.editor import VideoFileClip
from datetime import datetime

def convert_seconds_to_hms(seconds):
    """Convert seconds to hours, minutes, and seconds."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return hours, minutes, seconds

def format_duration(seconds):
    """Format the duration into a string 'H:M:S'."""
    hours, minutes, seconds = convert_seconds_to_hms(seconds)
    return f"{hours}h {minutes}m {seconds}s"

def get_timestamp():
    """Get the current date and time as a string."""
    now = datetime.now()
    return now.strftime('%Y%m%d_%H%M%S')

def get_file_size_in_kb(file_path):
    """Return the size of the file in kilobytes (KB)."""
    size_in_bytes = os.path.getsize(file_path)
    size_in_kb = size_in_bytes / 1024  # Convert bytes to KB
    return round(size_in_kb, 2)

# Ask user for the folder path containing the video files
folder_path = input("Please Enter Folder Path : ")

# Check if the provided path is valid
if not os.path.isdir(folder_path):
    print(f"The provided path '{folder_path}' is not a valid directory.")
    exit()

# List to store video details
video_data = []

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    # Check if the file is a video (you can add more formats if needed)
    if filename.endswith(('.mp4', '.avi', '.mkv', '.mov')):
        file_path = os.path.join(folder_path, filename)
        
        # Get the duration, resolution, and size of the video
        try:
            with VideoFileClip(file_path) as video:
                duration = video.duration  # Duration in seconds
                resolution = video.size    # Resolution as [width, height]
        except Exception as e:
            print(f"Error processing file '{filename}': {e}")
            continue
        
        # Get the file size in KB
        file_size_kb = get_file_size_in_kb(file_path)
        
        # Format the duration
        formatted_duration = format_duration(duration)

        # Extract the file extension (type) and remove the leading dot
        file_extension = os.path.splitext(filename)[1][1:]  # Extracts the part after the last dot
        
        # Append video name, duration, resolution, file size, and type to the list
        video_data.append({
            'Video Name': filename,
            'Duration (H:M:S)': formatted_duration,
            'Resolution': f"{resolution[0]}x{resolution[1]}",  # Format resolution as 'widthxheight'
            'File Type': file_extension,
            'File Size (KB)': file_size_kb,
        })

# Create a DataFrame from the video data
df = pd.DataFrame(video_data)

# Generate the output file name with current date and time
timestamp = get_timestamp()
output_excel_path = f'video_details_{timestamp}.xlsx'
df.to_excel(output_excel_path, index=False)

# Find duplicates based on 'Duration', 'Resolution', 'File Size', and 'File Type'
duplicates = df[df.duplicated(subset=['Duration (H:M:S)', 'Resolution', 'File Size (KB)', 'File Type'], keep=False)]

# Check if duplicates were found and display them
if not duplicates.empty:
    print("Duplicate video files found:")
    print("--------------------------------")
    print(duplicates[['Video Name', 'Duration (H:M:S)', 'Resolution', 'File Size (KB)', 'File Type']])
    print("--------------------------------")
else:
    print("No duplicate video files found.")

# Save the duplicates to a separate Excel file (optional)
if not duplicates.empty:
    duplicates_output_excel_path = f'duplicate_videos_{timestamp}.xlsx'
    duplicates.to_excel(duplicates_output_excel_path, index=False)
    print(f"Duplicate video details exported to {duplicates_output_excel_path}")
