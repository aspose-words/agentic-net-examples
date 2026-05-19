using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentWithCustomNaming
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections to demonstrate splitting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - First paragraph.");

        // Configure HTML save options to split by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Base file name for the main HTML file and for split parts.
        string baseFileName = "SplitDocument.html";

        // Assign the custom callback that renames each document part.
        saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(outputDir, baseFileName, saveOptions.DocumentSplitCriteria);

        // Save the document; this will produce multiple HTML files.
        string mainOutputPath = Path.Combine(outputDir, baseFileName);
        doc.Save(mainOutputPath, saveOptions);

        // Verify that split files were created.
        var partFiles = Directory.GetFiles(outputDir)
                                 .Where(f => Path.GetFileName(f).StartsWith(Path.GetFileNameWithoutExtension(baseFileName) + " part"))
                                 .ToArray();

        if (partFiles.Length == 0)
            throw new InvalidOperationException("No split document parts were generated.");

        Console.WriteLine($"Main document saved to: {mainOutputPath}");
        Console.WriteLine("Generated split parts:");
        foreach (var file in partFiles)
            Console.WriteLine($" - {file}");
    }

    // Callback that customizes the file name and stream for each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partCounter = 0;

        public SavedDocumentPartRename(string outputDir, string baseFileName, DocumentSplitCriteria criteria)
        {
            _outputDir = outputDir;
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the type of part based on the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)} part {++_partCounter}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (without path) and the full stream where Aspose.Words will write the part.
            args.DocumentPartFileName = partFileName;
            string fullPath = Path.Combine(_outputDir, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
