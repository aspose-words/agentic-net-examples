using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX file with several pages.
        string docPath = Path.Combine(workDir, "Sample.docx");
        CreateSampleDocument(docPath);

        // List of documents to process (batch). Here we have only one, but the code works for many.
        List<string> documents = new List<string> { docPath };

        // Process the batch in parallel.
        Parallel.ForEach(documents, docFile =>
        {
            // Load the document.
            Document doc = new Document(docFile);

            // Configure image save options for TIFF with multi‑frame layout.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Ensure each page becomes a separate frame.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Optional: set resolution for better quality.
                Resolution = 300
            };

            // Determine output file name.
            string tiffPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docFile) + ".tiff");

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, options);
        });

        // Validate that the TIFF file was created.
        string resultTiff = Path.Combine(outputDir, "Sample.tiff");
        if (!File.Exists(resultTiff) || new FileInfo(resultTiff).Length == 0)
            throw new InvalidOperationException("Failed to create the multipage TIFF file.");

        // Optionally, report success.
        Console.WriteLine("Multipage TIFF generated at: " + resultTiff);
    }

    // Creates a simple DOCX with three pages.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1: Introduction");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2: Content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3: Conclusion");

        doc.Save(filePath);
    }
}
