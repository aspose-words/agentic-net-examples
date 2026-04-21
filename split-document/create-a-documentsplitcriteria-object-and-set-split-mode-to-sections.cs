using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Create a DocumentSplitCriteria instance and set it to split by sections.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.SectionBreak;

        // Configure HtmlSaveOptions to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria,
            DocumentPartSavingCallback = new SavedDocumentPartRename(outputDir, "SplitDocument.html")
        };

        // Save the document; it will be split into separate HTML files, one per section.
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that split parts were created.
        var partFiles = Directory.GetFiles(outputDir, "SplitDocument_part_*.html");
        if (partFiles.Length < 3)
            throw new InvalidOperationException("Expected at least three split parts, but fewer were found.");

        // Optional: display the generated part file names.
        foreach (var file in partFiles)
            Console.WriteLine($"Created split part: {Path.GetFileName(file)}");
    }

    // Callback to customize the naming of each split document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string outputDir, string baseFileName)
        {
            _outputDir = outputDir;
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part_{_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, partFileName), FileMode.Create);
        }
    }
}
