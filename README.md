# PowerPoint-Accessibility
This program is for a Streamlit app to add alt text to PowerPoint files, with options to output pdf's of slides and handouts. This app uses an API call to the TAMU AI chat platform. Users enter their API key, and then select a preferred LLM for generate alt text. If the model supports image analysis, image content will be analyzed. Otherwise, slide text is used to generate alt text. The user can uploaded one or more PowerPoint files to process. Updated PowerPoint files can be downloaded. PDF's of slides or handouts can be optionally created using Windows COM automation. Output files are downloaded, zip files are created if multiple input files are used. User instructions:
1. Enter API key
2. Select desired model
3. (Optional) Choose output as slides or handouts
4. Upload one or more PowerPoint files
5. Process data
6. Download output files
