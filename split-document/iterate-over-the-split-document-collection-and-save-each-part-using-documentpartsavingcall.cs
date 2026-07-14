using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Define the base name for the split HTML files.
        string baseFileName = "SplitDocument.html";

        // Configure HTML save options to split by section and assign the callback.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(artifactsDir, baseFileName, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; the callback will be invoked for each part.
        doc.Save(Path.Combine(artifactsDir, baseFileName), saveOptions);

        // Validate that the expected number of parts were created.
        int expectedParts = doc.Sections.Count;
        string[] partFiles = Directory.GetFiles(artifactsDir, $"{Path.GetFileNameWithoutExtension(baseFileName)} part *.html");
        if (partFiles.Length != expectedParts)
            throw new InvalidOperationException($"Expected {expectedParts} parts, but found {partFiles.Length}.");

        // Optionally, display the created file names.
        foreach (string file in partFiles)
            Console.WriteLine("Created: " + Path.GetFileName(file));
    }

    // Callback that customizes the file name and stream for each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex;

        public SavedDocumentPartRename(string outputFolder, string baseFileName, DocumentSplitCriteria criteria)
        {
            _outputFolder = outputFolder;
            _baseFileName = Path.GetFileNameWithoutExtension(baseFileName);
            _criteria = criteria;
            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the type of split part for naming (section in this example).
            string partType = _criteria switch
            {
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"{_baseFileName} part {++_partIndex}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (without path) and provide a stream to write the part.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
