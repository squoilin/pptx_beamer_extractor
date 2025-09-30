# PPTX to Beamer Extractor

This script converts a PowerPoint (.pptx) presentation into a LaTeX Beamer presentation, extracting text and images from the original file.

## Features

- Extracts all text from each slide.
- Extracts all images from each slide.
- Creates a Beamer .tex file with one frame per slide.
- Saves images in a separate images folder.
- Attempts to extract presentation metadata (title, author).
- Sanitizes slide titles for use in image filenames.

## Installation

1. Clone the repository:
   git clone git@github.com:squoilin/pptx_beamer_extractor.git
   cd pptx_beamer_extractor

2. Create a Conda environment:
   conda create --name pptx_latex_extractor python=3.10 -y
   conda activate pptx_latex_extractor

3. Install dependencies:
   pip install python-pptx lxml

## Usage

Run the script from the command line, providing the input .pptx file and an output directory:

python pptx_to_beamer.py <input.pptx> <output_directory>

Alternatively, you can use the extract_pptx script to automate the entire process:

./extract_pptx <input.pptx>

### Example

An example presentation is provided in the example/ directory.

./extract_pptx example/nW_BE_scenario_25_09_27.pptx

This will create a nW_BE_scenario_25_09_27 directory with the final PDF and an images subdirectory.

To compile the LaTeX file, you will need a TeX distribution installed (e.g., TeX Live, MiKTeX).

pdflatex -output-directory=output_dir presentation.tex
