using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");

        // Configure HTML save options to split by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(outputDir, DocumentSplitCriteria.SectionBreak)
        };

        // Main output file name (the first part will be saved with this name if not overridden).
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that the expected parts were saved.
        string[] partFiles = Directory.GetFiles(outputDir, "SplitDocument_part_*.html");
        if (partFiles.Length != doc.Sections.Count)
            throw new InvalidOperationException($"Expected {doc.Sections.Count} parts, but found {partFiles.Length}.");

        // Example: list saved part files (no console output required by the task).
        // foreach (var file in partFiles) { /* process if needed */ }
    }

    // Callback that assigns a custom file name and stream for each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public SavedDocumentPartRename(string outputDir, DocumentSplitCriteria criteria)
        {
            _outputDir = outputDir;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a unique file name for each part.
            string partFileName = $"SplitDocument_part_{++_count}.html";

            // Set the file name (without path) and provide a stream to write the part.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, partFileName), FileMode.Create);
            // KeepDocumentPartStreamOpen remains false (default), so Aspose.Words will close the stream after saving.
        }
    }
}
