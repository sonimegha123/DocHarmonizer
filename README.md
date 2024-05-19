# DocHarmonizer  
Document Harmonizer is a Python project designed to convert DOCX documents into a consistent format. It extracts tables and images, converts the document to Markdown for processing, and then reconstructs the document in DOCX format while ensuring uniformity in style and format.  

# Features  
1. Extracts tables and images from DOCX documents.  
2. Converts DOCX to Markdown, allowing for easy text processing.  
3. Utilizes OpenAI's GPT-4 model to ensure consistent document style and format.  
4. Reconstructs the document in DOCX format, embedding extracted images and maintaining table placeholders.

# Usage  
1. Set up your environment variables. Create a .env file in the root directory and add your OpenAI API key and pandoc path:
   OPENAI_API_KEY=your_openai_api_key  
   PANDOC_PATH=/path/to/your/pandoc  

2. Prepare your DOCX file and specify the paths for the input and output files in the main function call:  
  if __name__ == "__main__":
    docx_file_path = "/path/to/your/input.docx"
    output_docx_path = "/path/to/your/output.docx"
    main(docx_file_path, output_docx_path)

3. Run the script:
   python your_script_name.py

