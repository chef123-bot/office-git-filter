#!/usr/bin/env python3
"""
office2text.py - Convert office documents to plain text for Git diff visualization.
Supports: DOCX, XLSX, PDF, PPTX, and other common office formats.
"""

import os
import sys
import subprocess
import tempfile
import json
import re
from pathlib import Path
from typing import Optional, List, Dict, Any
import argparse

# Optional imports for Python-based parsing
try:
    import docx2txt
    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False

try:
    import PyPDF2
    HAVE_PYPDF2 = True
except ImportError:
    HAVE_PYPDF2 = False

try:
    import pdfplumber
    HAVE_PDFPLUMBER = True
except ImportError:
    HAVE_PDFPLUMBER = False

try:
    import pandas as pd
    HAVE_PANDAS = True
except ImportError:
    HAVE_PANDAS = False

try:
    import openpyxl
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False

try:
    from pptx import Presentation
    HAVE_PPTX = True
except ImportError:
    HAVE_PPTX = False

try:
    import odf.opendocument
    import odf.text
    HAVE_ODF = True
except ImportError:
    HAVE_ODF = False


class OfficeToTextConverter:
    """Convert various office document formats to plain text."""
    
    def __init__(self, use_python_libs: bool = True, verbose: bool = False):
        """
        Initialize converter.
        
        Args:
            use_python_libs: Prefer Python libraries over external tools
            verbose: Enable verbose output for debugging
        """
        self.use_python_libs = use_python_libs
        self.verbose = verbose
        self.temp_dir = tempfile.mkdtemp(prefix="office2text_")
        
        # External tools to try (in order of preference)
        self.external_tools = {
            'pdf': ['pdftotext', 'mutool'],
            'docx': ['pandoc', 'unoconv', 'soffice'],
            'doc': ['antiword', 'catdoc', 'unoconv', 'soffice'],
            'xlsx': ['xlsx2csv', 'in2csv', 'ssconvert', 'unoconv', 'soffice'],
            'pptx': ['unoconv', 'soffice'],
            'odt': ['pandoc', 'unoconv', 'soffice'],
            'rtf': ['unrtf', 'pandoc', 'unoconv', 'soffice']
        }
        
        # Check which external tools are available
        self.available_tools = {}
        for format_type, tools in self.external_tools.items():
            self.available_tools[format_type] = []
            for tool in tools:
                if self._check_tool_available(tool):
                    self.available_tools[format_type].append(tool)
    
    def _check_tool_available(self, tool_name: str) -> bool:
        """Check if an external tool is available in PATH."""
        try:
            subprocess.run(
                ['which', tool_name] if sys.platform != 'win32' else ['where', tool_name],
                capture_output=True,
                check=False
            )
            return True
        except Exception:
            return False
    
    def _log(self, message: str, level: str = "INFO"):
        """Log messages if verbose mode is enabled."""
        if self.verbose:
            print(f"[{level}] {message}", file=sys.stderr)
    
    def convert(self, file_path: str) -> str:
        """
        Convert a file to plain text.
        
        Args:
            file_path: Path to the file to convert
            
        Returns:
            Plain text content of the file
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        ext = path.suffix.lower().lstrip('.')
        
        # Map of file extensions to conversion methods
        converters = {
            'pdf': self._convert_pdf,
            'docx': self._convert_docx,
            'doc': self._convert_doc,
            'xlsx': self._convert_xlsx,
            'xls': self._convert_xls,
            'pptx': self._convert_pptx,
            'ppt': self._convert_ppt,
            'odt': self._convert_odt,
            'ods': self._convert_ods,
            'odp': self._convert_odp,
            'rtf': self._convert_rtf,
            'txt': self._convert_text,
            'md': self._convert_text,
            'csv': self._convert_text,
            'tsv': self._convert_text,
        }
        
        converter = converters.get(ext)
        if converter:
            return converter(str(path))
        else:
            # Try to determine file type and convert
            return self._convert_unknown(str(path))
    
    def _convert_pdf(self, file_path: str) -> str:
        """Convert PDF to text."""
        # Try Python libraries first if requested
        if self.use_python_libs:
            if HAVE_PDFPLUMBER:
                try:
                    self._log(f"Using pdfplumber for {file_path}")
                    text = []
                    with pdfplumber.open(file_path) as pdf:
                        for page in pdf.pages:
                            page_text = page.extract_text()
                            if page_text:
                                text.append(page_text)
                    return "\n".join(text)
                except Exception as e:
                    self._log(f"pdfplumber failed: {e}", "WARNING")
            
            if HAVE_PYPDF2:
                try:
                    self._log(f"Using PyPDF2 for {file_path}")
                    text = []
                    with open(file_path, 'rb') as f:
                        pdf_reader = PyPDF2.PdfReader(f)
                        for page in pdf_reader.pages:
                            page_text = page.extract_text()
                            if page_text:
                                text.append(page_text)
                    return "\n".join(text)
                except Exception as e:
                    self._log(f"PyPDF2 failed: {e}", "WARNING")
        
        # Fall back to external tools
        for tool in self.available_tools.get('pdf', []):
            try:
                if tool == 'pdftotext':
                    self._log(f"Using pdftotext for {file_path}")
                    result = subprocess.run(
                        ['pdftotext', '-layout', '-nopgbrk', '-q', file_path, '-'],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'mutool':
                    self._log(f"Using mutool for {file_path}")
                    result = subprocess.run(
                        ['mutool', 'draw', '-F', 'text', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
            except subprocess.CalledProcessError as e:
                self._log(f"{tool} failed: {e}", "WARNING")
                continue
        
        raise RuntimeError(f"Failed to convert PDF: {file_path}. Install poppler-utils or mupdf-tools.")
    
    def _convert_docx(self, file_path: str) -> str:
        """Convert DOCX to text."""
        # Try Python library first
        if self.use_python_libs and HAVE_DOCX:
            try:
                self._log(f"Using docx2txt for {file_path}")
                return docx2txt.process(file_path)
            except Exception as e:
                self._log(f"docx2txt failed: {e}", "WARNING")
        
        # Fall back to external tools
        for tool in self.available_tools.get('docx', []):
            try:
                if tool == 'pandoc':
                    self._log(f"Using pandoc for {file_path}")
                    result = subprocess.run(
                        ['pandoc', '-f', 'docx', '-t', 'plain', '--wrap=none', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'unoconv':
                    self._log(f"Using unoconv for {file_path}")
                    result = subprocess.run(
                        ['unoconv', '--stdout', '-f', 'txt', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'soffice':
                    self._log(f"Using soffice for {file_path}")
                    temp_output = os.path.join(self.temp_dir, "output.txt")
                    subprocess.run(
                        ['soffice', '--headless', '--convert-to', 'txt:Text',
                         '--outdir', self.temp_dir, file_path],
                        capture_output=True,
                        check=True
                    )
                    output_file = os.path.join(self.temp_dir, 
                                             os.path.basename(file_path).rsplit('.', 1)[0] + '.txt')
                    if os.path.exists(output_file):
                        with open(output_file, 'r', encoding='utf-8', errors='ignore') as f:
                            return f.read()
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self._log(f"{tool} failed: {e}", "WARNING")
                continue
        
        raise RuntimeError(f"Failed to convert DOCX: {file_path}. Install docx2txt, pandoc, or libreoffice.")
    
    def _convert_doc(self, file_path: str) -> str:
        """Convert legacy DOC format to text."""
        # Try external tools
        for tool in self.available_tools.get('doc', []):
            try:
                if tool == 'antiword':
                    self._log(f"Using antiword for {file_path}")
                    result = subprocess.run(
                        ['antiword', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'catdoc':
                    self._log(f"Using catdoc for {file_path}")
                    result = subprocess.run(
                        ['catdoc', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'unoconv':
                    self._log(f"Using unoconv for {file_path}")
                    result = subprocess.run(
                        ['unoconv', '--stdout', '-f', 'txt', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'soffice':
                    self._log(f"Using soffice for {file_path}")
                    temp_output = os.path.join(self.temp_dir, "output.txt")
                    subprocess.run(
                        ['soffice', '--headless', '--convert-to', 'txt:Text',
                         '--outdir', self.temp_dir, file_path],
                        capture_output=True,
                        check=True
                    )
                    output_file = os.path.join(self.temp_dir, 
                                             os.path.basename(file_path).rsplit('.', 1)[0] + '.txt')
                    if os.path.exists(output_file):
                        with open(output_file, 'r', encoding='utf-8', errors='ignore') as f:
                            return f.read()
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self._log(f"{tool} failed: {e}", "WARNING")
                continue
        
        raise RuntimeError(f"Failed to convert DOC: {file_path}. Install antiword, catdoc, or libreoffice.")
    
    def _convert_xlsx(self, file_path: str) -> str:
        """Convert XLSX to CSV-like text."""
        # Try Python libraries first
        if self.use_python_libs:
            if HAVE_PANDAS:
                try:
                    self._log(f"Using pandas for {file_path}")
                    # Read all sheets
                    excel_file = pd.ExcelFile(file_path)
                    output = []
                    
                    for sheet_name in excel_file.sheet_names:
                        df = excel_file.parse(sheet_name)
                        output.append(f"=== Sheet: {sheet_name} ===")
                        output.append(df.to_csv(sep='|', index=False))
                    
                    return "\n".join(output)
                except Exception as e:
                    self._log(f"pandas failed: {e}", "WARNING")
            
            elif HAVE_OPENPYXL:
                try:
                    self._log(f"Using openpyxl for {file_path}")
                    from openpyxl import load_workbook
                    
                    wb = load_workbook(filename=file_path, read_only=True, data_only=True)
                    output = []
                    
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        output.append(f"=== Sheet: {sheet_name} ===")
                        
                        # Get dimensions
                        max_row = ws.max_row
                        max_column = ws.max_column
                        
                        # Convert column number to letter
                        from openpyxl.utils import get_column_letter
                        
                        for row in ws.iter_rows(min_row=1, max_row=max_row, 
                                               min_col=1, max_col=max_column):
                            row_values = []
                            for cell in row:
                                value = cell.value
                                if value is None:
                                    value = ""
                                # Escape pipe characters
                                if isinstance(value, str):
                                    value = value.replace('|', '\\|')
                                row_values.append(str(value))
                            output.append("|".join(row_values))
                    
                    return "\n".join(output)
                except Exception as e:
                    self._log(f"openpyxl failed: {e}", "WARNING")
        
        # Fall back to external tools
        for tool in self.available_tools.get('xlsx', []):
            try:
                if tool == 'xlsx2csv':
                    self._log(f"Using xlsx2csv for {file_path}")
                    result = subprocess.run(
                        ['xlsx2csv', '-d', '|', '-a', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'in2csv':
                    self._log(f"Using in2csv for {file_path}")
                    result = subprocess.run(
                        ['in2csv', '-f', 'xlsx', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'ssconvert':
                    self._log(f"Using ssconvert for {file_path}")
                    temp_output = os.path.join(self.temp_dir, "output.csv")
                    subprocess.run(
                        ['ssconvert', '-T', 'Gnumeric_stf:stf_csv', 
                         '--export-type=Gnumeric_stf:stf_csv', file_path, temp_output],
                        capture_output=True,
                        check=True
                    )
                    if os.path.exists(temp_output):
                        with open(temp_output, 'r', encoding='utf-8') as f:
                            return f.read()
                elif tool in ['unoconv', 'soffice']:
                    # Similar to DOC conversion but for CSV
                    self._log(f"Using {tool} for {file_path}")
                    temp_output = os.path.join(self.temp_dir, "output.csv")
                    if tool == 'unoconv':
                        result = subprocess.run(
                            ['unoconv', '--stdout', '-f', 'csv', file_path],
                            capture_output=True,
                            text=True,
                            check=True
                        )
                        return result.stdout
                    else:
                        subprocess.run(
                            ['soffice', '--headless', '--convert-to', 'csv:Text -txt-44:44,34,76',
                             '--outdir', self.temp_dir, file_path],
                            capture_output=True,
                            check=True
                        )
                        output_file = os.path.join(self.temp_dir, 
                                                 os.path.basename(file_path).rsplit('.', 1)[0] + '.csv')
                        if os.path.exists(output_file):
                            with open(output_file, 'r', encoding='utf-8') as f:
                                return f.read()
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self._log(f"{tool} failed: {e}", "WARNING")
                continue
        
        raise RuntimeError(f"Failed to convert XLSX: {file_path}. Install pandas, xlsx2csv, or libreoffice.")
    
    def _convert_xls(self, file_path: str) -> str:
        """Convert legacy XLS format to text."""
        # Similar to XLSX but for older format
        if self.use_python_libs and HAVE_PANDAS:
            try:
                self._log(f"Using pandas for {file_path}")
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
                output = []
                
                for sheet_name in excel_file.sheet_names:
                    df = excel_file.parse(sheet_name)
                    output.append(f"=== Sheet: {sheet_name} ===")
                    output.append(df.to_csv(sep='|', index=False))
                
                return "\n".join(output)
            except Exception as e:
                self._log(f"pandas/xlrd failed: {e}", "WARNING")
        
        # Try external tools (same as xlsx)
        return self._convert_xlsx(file_path)
    
    def _convert_pptx(self, file_path: str) -> str:
        """Convert PPTX to text."""
        # Try Python library first
        if self.use_python_libs and HAVE_PPTX:
            try:
                self._log(f"Using python-pptx for {file_path}")
                prs = Presentation(file_path)
                text_runs = []
                
                for slide_number, slide in enumerate(prs.slides):
                    text_runs.append(f"=== Slide {slide_number + 1} ===")
                    
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            text = shape.text.strip()
                            if text:
                                text_runs.append(text)
                
                return "\n".join(text_runs)
            except Exception as e:
                self._log(f"python-pptx failed: {e}", "WARNING")
        
        # Fall back to external tools
        for tool in self.available_tools.get('pptx', []):
            try:
                if tool == 'unoconv':
                    self._log(f"Using unoconv for {file_path}")
                    result = subprocess.run(
                        ['unoconv', '--stdout', '-f', 'txt', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'soffice':
                    self._log(f"Using soffice for {file_path}")
                    temp_output = os.path.join(self.temp_dir, "output.txt")
                    subprocess.run(
                        ['soffice', '--headless', '--convert-to', 'txt:Text',
                         '--outdir', self.temp_dir, file_path],
                        capture_output=True,
                        check=True
                    )
                    output_file = os.path.join(self.temp_dir, 
                                             os.path.basename(file_path).rsplit('.', 1)[0] + '.txt')
                    if os.path.exists(output_file):
                        with open(output_file, 'r', encoding='utf-8', errors='ignore') as f:
                            return f.read()
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self._log(f"{tool} failed: {e}", "WARNING")
                continue
        
        raise RuntimeError(f"Failed to convert PPTX: {file_path}. Install python-pptx or libreoffice.")
    
    def _convert_ppt(self, file_path: str) -> str:
        """Convert legacy PPT format to text."""
        # Use same external tools as pptx
        return self._convert_pptx(file_path)
    
    def _convert_odt(self, file_path: str) -> str:
        """Convert ODT (OpenDocument Text) to text."""
        # Try Python library first
        if self.use_python_libs and HAVE_ODF:
            try:
                self._log(f"Using odf library for {file_path}")
                doc = odf.opendocument.load(file_path)
                text_content = []
                
                # Extract text from paragraphs
                for elem in doc.getElementsByType(odf.text.P):
                    text_content.append(elem)
                
                return "\n".join(str(t) for t in text_content)
            except Exception as e:
                self._log(f"odf library failed: {e}", "WARNING")
        
        # Try external tools (same as docx)
        return self._convert_docx(file_path)
    
    def _convert_ods(self, file_path: str) -> str:
        """Convert ODS (OpenDocument Spreadsheet) to text."""
        # Similar to xlsx conversion
        return self._convert_xlsx(file_path)
    
    def _convert_odp(self, file_path: str) -> str:
        """Convert ODP (OpenDocument Presentation) to text."""
        # Similar to pptx conversion
        return self._convert_pptx(file_path)
    
    def _convert_rtf(self, file_path: str) -> str:
        """Convert RTF to text."""
        # Try external tools
        for tool in self.available_tools.get('rtf', []):
            try:
                if tool == 'unrtf':
                    self._log(f"Using unrtf for {file_path}")
                    result = subprocess.run(
                        ['unrtf', '--text', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool == 'pandoc':
                    self._log(f"Using pandoc for {file_path}")
                    result = subprocess.run(
                        ['pandoc', '-f', 'rtf', '-t', 'plain', '--wrap=none', file_path],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    return result.stdout
                elif tool in ['unoconv', 'soffice']:
                    # Similar to doc conversion
                    return self._convert_doc(file_path)
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self._log(f"{tool} failed: {e}", "WARNING")
                continue
        
        raise RuntimeError(f"Failed to convert RTF: {file_path}. Install unrtf or pandoc.")
    
    def _convert_text(self, file_path: str) -> str:
        """Read plain text files."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except UnicodeDecodeError:
            # Try with different encoding
            with open(file_path, 'r', encoding='latin-1') as f:
                return f.read()
    
    def _convert_unknown(self, file_path: str) -> str:
        """Try to convert unknown file types."""
        # Try to determine file type
        try:
            import magic
            mime = magic.Magic(mime=True)
            mime_type = mime.from_file(file_path)
            
            if 'pdf' in mime_type:
                return self._convert_pdf(file_path)
            elif 'word' in mime_type or 'officedocument.wordprocessingml' in mime_type:
                return self._convert_docx(file_path)
            elif 'excel' in mime_type or 'officedocument.spreadsheetml' in mime_type:
                return self._convert_xlsx(file_path)
            elif 'powerpoint' in mime_type or 'officedocument.presentationml' in mime_type:
                return self._convert_pptx(file_path)
            elif 'opendocument' in mime_type:
                if 'text' in mime_type:
                    return self._convert_odt(file_path)
                elif 'spreadsheet' in mime_type:
                    return self._convert_ods(file_path)
                elif 'presentation' in mime_type:
                    return self._convert_odp(file_path)
            elif 'text' in mime_type:
                return self._convert_text(file_path)
        except ImportError:
            # python-magic not available
            pass
        
        # As last resort, try to read as text
        try:
            return self._convert_text(file_path)
        except:
            raise RuntimeError(f"Unsupported file type: {file_path}")
    
    def cleanup(self):
        """Clean up temporary files."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description='Convert office documents to plain text for Git diff',
        epilog='Example: office2text.py document.docx'
    )
    parser.add_argument('file', help='File to convert')
    parser.add_argument('--no-python', action='store_true',
                       help='Do not use Python libraries, only external tools')
    parser.add_argument('--list-tools', action='store_true',
                       help='List available conversion tools')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Verbose output')
    parser.add_argument('--output', '-o', help='Output file (default: stdout)')
    
    args = parser.parse_args()
    
    if args.list_tools:
        converter = OfficeToTextConverter(verbose=args.verbose)
        print("Available external tools:")
        for fmt, tools in converter.available_tools.items():
            if tools:
                print(f"  {fmt}: {', '.join(tools)}")
            else:
                print(f"  {fmt}: None available")
        
        print("\nPython libraries:")
        libs = []
        if HAVE_DOCX: libs.append("docx2txt")
        if HAVE_PYPDF2: libs.append("PyPDF2")
        if HAVE_PDFPLUMBER: libs.append("pdfplumber")
        if HAVE_PANDAS: libs.append("pandas")
        if HAVE_OPENPYXL: libs.append("openpyxl")
        if HAVE_PPTX: libs.append("python-pptx")
        if HAVE_ODF: libs.append("odf")
        print(f"  Installed: {', '.join(libs) if libs else 'None'}")
        return 0
    
    # Convert the file
    converter = OfficeToTextConverter(
        use_python_libs=not args.no_python,
        verbose=args.verbose
    )
    
    try:
        text = converter.convert(args.file)
        
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            if args.verbose:
                print(f"Output written to: {args.output}", file=sys.stderr)
        else:
            print(text)
        
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    finally:
        converter.cleanup()


if __name__ == '__main__':
    sys.exit(main())