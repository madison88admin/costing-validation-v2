#!/bin/bash
# Create a dist folder for deployment
mkdir -p dist

# Copy all files from src to dist
cp -r src/* dist/

# Copy public folder contents to dist
cp -r public/* dist/

echo "Build complete! Files ready in dist folder"
