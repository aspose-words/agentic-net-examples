using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections to demonstrate splitting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Configure HTML save options to split the document by section break.
        string mainFileName = "SplitDocument.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(outputDir, mainFileName, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; the callback will be invoked for each part.
        doc.Save(Path.Combine(outputDir, mainFileName), saveOptions);

        // Validate that the expected number of parts were created.
        string partPattern = $"{Path.GetFileNameWithoutExtension(mainFileName)} part *.html";
        string[] partFiles = Directory.GetFiles(outputDir, partPattern);
        if (partFiles.Length != doc.Sections.Count)
            throw new Exception($"Expected {doc.Sections.Count} parts, but found {partFiles.Length}.");

        // Example completed without interactive prompts.
    }

    // Callback that assigns a custom file name and stream for each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex;

        public SavedDocumentPartRename(string outputDir, string baseFileName, DocumentSplitCriteria criteria)
        {
            _outputDir = outputDir;
            _baseFileName = baseFileName;
            _criteria = criteria;
            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Build a deterministic part file name.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)} part {_partIndex + 1}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name and provide a stream where Aspose.Words will write the part.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, partFileName), FileMode.Create);

            _partIndex++;
        }
    }
}
