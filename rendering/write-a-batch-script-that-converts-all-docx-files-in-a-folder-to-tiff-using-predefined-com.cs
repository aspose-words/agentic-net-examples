using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffDemo");
        Directory.CreateDirectory(workFolder);

        // Create sample DOCX files in the folder.
        CreateSampleDocx(Path.Combine(workFolder, "Sample1.docx"), "First document.\nPage 2.", 2);
        CreateSampleDocx(Path.Combine(workFolder, "Sample2.docx"), "Second document with more content.\nAnother page.\nThird page.", 3);

        // Convert each DOCX file in the folder to a TIFF image with LZW compression.
        foreach (string docxPath in Directory.GetFiles(workFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Prepare TIFF save options with LZW compression.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use LZW compression for the TIFF output.
                TiffCompression = TiffCompression.Lzw
                // By default all pages are rendered; no need to set PageSet.
            };

            // Determine output TIFF file path.
            string tiffPath = Path.ChangeExtension(docxPath, ".tiff");

            // Save the document as TIFF.
            doc.Save(tiffPath, options);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: clean up the temporary folder after processing.
        // Directory.Delete(workFolder, true);
    }

    // Helper method to create a simple DOCX file with given text and a specified number of pages.
    private static void CreateSampleDocx(string filePath, string content, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the provided content.
        builder.Writeln(content);

        // Insert manual page breaks to reach the desired page count.
        for (int i = 1; i < pageCount; i++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page {i + 1}");
        }

        // Save the sample DOCX.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
