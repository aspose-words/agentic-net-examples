using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Saving; // For MultiPageLayout

class Program
{
    static void Main()
    {
        // Input DOCX files to be converted.
        List<string> docxFiles = new List<string>
        {
            @"C:\Input\Document1.docx",
            @"C:\Input\Document2.docx",
            // Add more file paths as needed.
        };

        // Output directory for the generated multipage TIFF files.
        string outputDir = @"C:\Output\TIFFs";
        Directory.CreateDirectory(outputDir);

        // Process each DOCX file in parallel.
        Parallel.ForEach(docxFiles, docxPath =>
        {
            // Load the DOCX document (uses the Document(string) constructor rule).
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

            // Use the TiffFrames layout so each page becomes a separate frame
            // in the resulting multi‑frame TIFF (uses MultiPageLayout.TiffFrames rule).
            options.PageLayout = MultiPageLayout.TiffFrames();

            // Optional: set resolution or other image options here.
            // options.Resolution = 300;

            // Build the output file name.
            string fileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string outputPath = Path.Combine(outputDir, fileName);

            // Save the document as a multipage TIFF (uses Document.Save(string, SaveOptions) rule).
            doc.Save(outputPath, options);
        });
    }
}
