#!/bin/bash

# Source directory (gas-production)
SOURCE_DIR="gas-production"

# Destination directory (gas-development)
DEST_DIR="gas-development"

# Ensure the destination directory exists
mkdir -p "$DEST_DIR"

# Copy files from the source to the destination
cp -r "$SOURCE_DIR/"* "$DEST_DIR/"

echo "Files copied from $SOURCE_DIR to $DEST_DIR"
