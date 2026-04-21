using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with several sections to trigger splitting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 4; i++)
        {
            builder.Writeln($"This is content of section {i}.");
            // Insert a section break to create separate parts.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Configure HTML save options with section‑based splitting.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new PartSavingCallback(outputDir)
        };

        // Save the document; the callback will create separate files for each part.
        string mainFile = Path.Combine(outputDir, "Combined.html");
        doc.Save(mainFile, saveOptions);

        // Verify that the expected part files were created.
        var partFiles = Directory.GetFiles(outputDir)
                                 .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                                             f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                                 .OrderBy(f => f)
                                 .ToArray();

        if (partFiles.Length == 0)
            throw new InvalidOperationException("No document parts were generated.");

        Console.WriteLine("Generated document parts:");
        foreach (var file in partFiles)
            Console.WriteLine($" - {Path.GetFileName(file)}");
    }

    // Callback that assigns .docx to even parts and .pdf to odd parts.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private int _partIndex = 0;

        public PartSavingCallback(string outputDir)
        {
            _outputDir = outputDir;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _partIndex++;

            // Determine extension based on part index (odd -> .pdf, even -> .docx).
            string extension = (_partIndex % 2 == 0) ? ".docx" : ".pdf";

            // Build a unique file name for the part.
            string partFileName = $"Part_{_partIndex}{extension}";
            string fullPath = Path.Combine(_outputDir, partFileName);

            // Set the file name (without path) and the stream where Aspose.Words will write the part.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
