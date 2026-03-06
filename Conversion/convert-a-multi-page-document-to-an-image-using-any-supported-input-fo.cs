using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertMultiPageToImages
{
    static void Main()
    {
        // Load a multi‑page document from any supported format (e.g., DOCX, PDF, etc.).
        // Replace the file path with the actual location of your source document.
        string inputFile = "input.docx";
        Document doc = new Document(inputFile);

        // Configure image save options. Here we choose PNG format and a high resolution.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);
        imgOptions.Resolution = 300; // DPI

        // Iterate through all pages of the document.
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Render only the current page by setting the PageSet to the zero‑based page index.
            imgOptions.PageSet = new PageSet(i);

            // Define the output file name for the current page image.
            string outputFile = $"Page_{i + 1}.png";

            // Save the rendered page image using the Document.Save method.
            doc.Save(outputFile, imgOptions);
        }
    }
}
