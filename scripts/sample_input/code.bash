#!/bin/bash

# Define source and backup directories
SOURCE="/path/to/source"
BACKUP="/path/to/backup"

# Function to perform file sync using rsync
sync_files() {
    rsync -av --delete "$SOURCE" "$BACKUP"
}

# Log start of the script
echo "Starting backup: $(date)"

# Call the file sync function
sync_files

# Check if rsync was successful
if [ $? -eq 0 ]; then
    echo "Backup completed successfully."
else
    echo "Backup failed with error code $?."
fi

# Log end of the script
echo "Backup process finished: $(date)"