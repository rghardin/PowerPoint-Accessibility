# -*- coding: utf-8 -*-
"""
Created on Wed Apr  1 05:42:01 2026

@author: robert.hardin

pdf_converter.py
Standalone script to convert PowerPoint files to PDF using Windows COM automation.
Called via subprocess from the main Streamlit application.

Usage:
    python pdf_converter.py <input_pptx> <output_dir> <generate_slides> <generate_handouts> <slides_per_page>

Arguments:
    input_pptx: Path to the input PowerPoint file
    output_dir: Directory to save the output PDF files
    generate_slides: "true" or "false" - whether to generate slides PDF
    generate_handouts: "true" or "false" - whether to generate handouts PDF
    slides_per_page: Number of slides per page for handouts (2, 3, 4, or 6)

Output:
    Creates PDF files in the output directory:
    - <filename>_slides.pdf (if generate_slides is true)
    - <filename>_handouts.pdf (if generate_handouts is true)
    
    Prints JSON result to stdout with status and file paths.
"""

import sys
import os
import json
import pythoncom
import win32com.client


def get_handout_output_type(slides_per_page):
    """
    Convert user selection of slides per page to PowerPoint output type constant.
    Slides = 1, 2 slides = 2, 3 slides = 3, 4 slides = 8, 6 slides = 4
    """
    output_type_map = {
        1: 1,   # Slides
        2: 2,   # 2 slides per page
        3: 3,   # 3 slides per page
        4: 8,   # 4 slides per page
        6: 4    # 6 slides per page
    }
    return output_type_map.get(slides_per_page, 3)


def convert_pptx_to_pdf(input_pptx, output_dir, generate_slides=True, generate_handouts=False, slides_per_page=3):
    """
    Convert a PowerPoint file to PDF using Windows COM automation.
    
    Args:
        input_pptx: Path to the input PowerPoint file
        output_dir: Directory to save the output PDF files
        generate_slides: Whether to generate slides PDF
        generate_handouts: Whether to generate handouts PDF
        slides_per_page: Number of slides per page for handouts
    
    Returns:
        Dictionary with status and output file paths
    """
    result = {
        "success": False,
        "slides_pdf": None,
        "handouts_pdf": None,
        "error": None
    }
    
    # Get base filename without extension
    base_name = os.path.splitext(os.path.basename(input_pptx))[0]
    
    # Get handout output type
    handout_output_type = get_handout_output_type(slides_per_page)
    
    app = None
    presentation = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Initialize PowerPoint application
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        
        # Open the presentation
        presentation = app.Presentations.Open(os.path.abspath(input_pptx))
        
        # Generate slides PDF if requested
        if generate_slides:
            slides_pdf_path = os.path.join(output_dir, f"{base_name}_slides.pdf")
            presentation.ExportAsFixedFormat(
                os.path.abspath(slides_pdf_path),
                2,  # ppFixedFormatTypePDF
                OutputType=1,  # Slides
                HandoutOrder=2,  # Horizontal
                PrintRange=None
            )
            result["slides_pdf"] = slides_pdf_path
        
        # Generate handouts PDF if requested
        if generate_handouts:
            handouts_pdf_path = os.path.join(output_dir, f"{base_name}_handouts.pdf")
            presentation.ExportAsFixedFormat(
                os.path.abspath(handouts_pdf_path),
                2,  # ppFixedFormatTypePDF
                OutputType=handout_output_type,
                HandoutOrder=2,  # Horizontal
                PrintRange=None
            )
            result["handouts_pdf"] = handouts_pdf_path
        
        result["success"] = True
        
    except Exception as e:
        result["error"] = str(e)
    
    finally:
        # Close the presentation
        if presentation:
            try:
                presentation.Close()
            except:
                pass
            del presentation
        
        # Quit PowerPoint
        if app:
            try:
                app.Quit()
            except:
                pass
            del app
        
        # Uninitialize COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    return result


def main():
    """Main entry point for command-line usage."""
    if len(sys.argv) < 6:
        print(json.dumps({
            "success": False,
            "error": "Usage: python pdf_converter.py <input_pptx> <output_dir> <generate_slides> <generate_handouts> <slides_per_page>"
        }))
        sys.exit(1)
    
    input_pptx = sys.argv[1]
    output_dir = sys.argv[2]
    generate_slides = sys.argv[3].lower() == "true"
    generate_handouts = sys.argv[4].lower() == "true"
    slides_per_page = int(sys.argv[5])
    
    # Validate input file exists
    if not os.path.exists(input_pptx):
        print(json.dumps({
            "success": False,
            "error": f"Input file not found: {input_pptx}"
        }))
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Convert the file
    result = convert_pptx_to_pdf(
        input_pptx,
        output_dir,
        generate_slides,
        generate_handouts,
        slides_per_page
    )
    
    # Output result as JSON
    print(json.dumps(result))
    
    # Exit with appropriate code
    sys.exit(0 if result["success"] else 1)


if __name__ == "__main__":
    main()