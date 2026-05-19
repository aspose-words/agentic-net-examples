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
        string inputDir = Path.Combine(workDir, "Input");
        string outputDir = Path.Combine(workDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file with several pages.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(inputDir, "Sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add three pages of text.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document.
        sampleDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 2. Prepare a list of source documents to be processed in parallel.
        //    (Here we simply reuse the same file multiple times to simulate a batch.)
        // -----------------------------------------------------------------
        List<string> sourceFiles = new List<string>
        {
            sourceDocPath,
            sourceDocPath,
            sourceDocPath,
            sourceDocPath
        };

        // -----------------------------------------------------------------
        // 3. Process each document in parallel, rendering a multipage TIFF.
        // -----------------------------------------------------------------
        Parallel.ForEach(sourceFiles, (srcPath, state, index) =>
        {
            // Load the document.
            Document doc = new Document(srcPath);

            // Configure image save options for a multipage TIFF.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Each page will be stored as a separate frame in the TIFF.
                PageLayout = MultiPageLayout.TiffFrames(),
                // Optional: set resolution (dpi) for better quality.
                Resolution = 300
            };

            // Build the output file name.
            string outPath = Path.Combine(outputDir, $"Result_{index}.tiff");

            // Save the document as a multipage TIFF.
            doc.Save(outPath, options);

            // Validate that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outPath}");

            // Additional validation: ensure the source document has pages.
            if (doc.PageCount == 0)
                throw new InvalidOperationException($"Source document has no pages: {srcPath}");
        });

        // -----------------------------------------------------------------
        // 4. Final verification – list generated files.
        // -----------------------------------------------------------------
        Console.WriteLine("Multipage TIFF files generated:");
        foreach (string file in Directory.GetFiles(outputDir, "*.tiff"))
        {
            FileInfo info = new FileInfo(file);
            Console.WriteLine($"{Path.GetFileName(file)} – {info.Length} bytes");
        }
    }
}
