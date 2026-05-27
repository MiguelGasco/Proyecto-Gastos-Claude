#!/bin/bash
# Copy static files to output directory
mkdir -p /vercel/output/static
cp index.html /vercel/output/index.html
cp dashboard.html /vercel/output/dashboard.html
cp investments.html /vercel/output/investments.html
echo "Build complete"
