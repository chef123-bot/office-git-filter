#!/bin/bash
# Test script for office2text converter

echo "========================================"
echo "Testing office2text Converter"
echo "========================================"

cd ~/test/office3/bin/test_files

# Test 1: Basic text file
echo -e "\nTest 1: Basic text file"
bin test_files/simple.txt

# Test 2: CSV file
echo -e "\nTest 2: CSV file"
bin test_files/sample.csv

# Test 3: Markdown file
echo -e "\nTest 3: Markdown file"
bin test_files/sample.md | head -10

# Test 4: Complex text
echo -e "\nTest 4: Complex text file"
bin test_files/complex_test.txt | head -15

# Test 5: Test help
echo -e "\nTest 5: Help message"
bin --help

# Test 6: List all test files
echo -e "\nTest 6: Available test files"
find test_files -type f | while read file; do
    echo "  - $file"
done

echo -e "\n========================================"
echo "To test with actual office files:"
echo "1. Place .docx, .xlsx, .pptx files in test_files/"
echo "2. Run: bin test_files/yourfile.docx"
echo "========================================"
