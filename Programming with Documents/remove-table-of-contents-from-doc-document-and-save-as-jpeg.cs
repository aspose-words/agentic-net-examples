using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class RemoveTocAndConvertToJpeg
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path to the output JPEG image.
        string outputPath = @"C:\Docs\SourceDocument_NoToc.jpg";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all Table of Contents (TOC) fields from the document.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Configure image save options for JPEG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: set JPEG quality (0‑100). Higher value = better quality.
            JpegQuality = 90,
            // Save only the first page (index 0).
            PageSet = new PageSet(0)
        };

        // Save the modified document as a JPEG image.
        doc.Save(outputPath, saveOptions);
    }
}
