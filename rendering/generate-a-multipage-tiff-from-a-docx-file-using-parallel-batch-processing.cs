using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with multiple pages.
        string docPath = "Sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");
        doc.Save(docPath);

        // Prepare output directory.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Documents to process (single file in this example).
        string[] sourceDocs = { docPath };

        // Parallel batch processing: convert each DOCX to a multipage TIFF.
        Parallel.ForEach(sourceDocs, sourcePath =>
        {
            // Load the source document.
            Document sourceDoc = new Document(sourcePath);

            // Configure image save options for a multipage TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageLayout = MultiPageLayout.TiffFrames(); // Each page as a separate TIFF frame.

            // Determine output file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(sourcePath) + ".tiff";
            string tiffPath = Path.Combine(outputDir, tiffFileName);

            // Save the document as a multipage TIFF.
            sourceDoc.Save(tiffPath, options);

            // Validate that the TIFF file was created and is not empty.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"TIFF file was not created: {tiffPath}");
            if (new FileInfo(tiffPath).Length == 0)
                throw new InvalidOperationException($"TIFF file is empty: {tiffPath}");
        });

        // Execution completes without interactive prompts.
    }
}
