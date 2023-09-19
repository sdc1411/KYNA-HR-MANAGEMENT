#!/bin/bash

# Source directory (gas-development)
SOURCE_DIR="gas-development"

# Destination directory (gas-staging)
DEST_DIR="gas-staging"

# Ensure the destination directory exists
mkdir -p "$DEST_DIR"

# Copy files from the source to the destination
cp -r "$SOURCE_DIR/"* "$DEST_DIR/"

echo "Files copied from $SOURCE_DIR to $DEST_DIR"
