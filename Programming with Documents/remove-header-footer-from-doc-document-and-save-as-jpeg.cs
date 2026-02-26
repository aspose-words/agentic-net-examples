using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeaderFooterAndSaveAsJpeg
{
    static void Main()
    {
        // Input DOC file path
        string inputPath = @"C:\Docs\input.doc";

        // Output JPEG file path
        string outputPath = @"C:\Docs\output.jpg";

        // Load the document (lifecycle rule: Document(string))
        Document doc = new Document(inputPath);

        // Remove all headers and footers from every section
        foreach (Section section in doc.Sections)
        {
            // Clear the HeadersFooters collection for the current section
            section.HeadersFooters.Clear();
        }

        // Prepare image save options for JPEG (lifecycle rule: ImageSaveOptions(SaveFormat))
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: set JPEG quality (0-100). 100 = best quality.
            JpegQuality = 100
        };

        // Save the document as a JPEG image (lifecycle rule: Document.Save(string, SaveOptions))
        doc.Save(outputPath, saveOptions);
    }
}
