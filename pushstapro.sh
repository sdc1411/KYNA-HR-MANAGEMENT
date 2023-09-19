#!/bin/bash

# Source directory (gas-staging)
SOURCE_DIR="gas-staging"

# Destination directory (gas-production)
DEST_DIR="gas-production"

# Ensure the destination directory exists
mkdir -p "$DEST_DIR"

# Copy files from the source to the destination
cp -r "$SOURCE_DIR/"* "$DEST_DIR/"

echo "Files copied from $SOURCE_DIR to $DEST_DIR"
