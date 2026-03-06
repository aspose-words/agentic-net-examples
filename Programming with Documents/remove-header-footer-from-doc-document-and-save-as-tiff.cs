using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndConvertToTiff
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting TIFF file will be saved.
        string outputPath = @"C:\Docs\ResultImage.tiff";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all header and footer contents from each section.
        foreach (Section section in doc.Sections)
        {
            // Clears the text of headers/footers but keeps the objects,
            // effectively making the document header/footer‑less.
            section.ClearHeadersFooters();
        }

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: choose compression (default is Lzw).
            // TiffCompression = TiffCompression.None,
            // Optional: set resolution if higher quality is required.
            // Resolution = 300
        };

        // Save the document as a TIFF image (multi‑page if the source has multiple pages).
        doc.Save(outputPath, saveOptions);
    }
}
