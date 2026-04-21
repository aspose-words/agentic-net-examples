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
        builder.Writeln("Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");

        // Configure HTML save options to split by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Base name for the output files.
        string baseFileName = "SplitDocument";

        // Assign the custom callback that will be invoked for each document part.
        saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(
            artifactsDir, baseFileName, saveOptions.DocumentSplitCriteria);

        // Save the document; the callback will create separate files for each part.
        string mainOutputPath = Path.Combine(artifactsDir, baseFileName + ".html");
        doc.Save(mainOutputPath, saveOptions);

        // Verify that the expected number of parts were saved (three sections → three parts).
        string[] partFiles = Directory.GetFiles(artifactsDir,
            $"{baseFileName} part *.html");
        if (partFiles.Length != doc.Sections.Count)
            throw new InvalidOperationException(
                $"Expected {doc.Sections.Count} parts, but found {partFiles.Length}.");

        // Optional: display the created file names.
        foreach (string file in partFiles)
            Console.WriteLine("Created part: " + Path.GetFileName(file));
    }

    // Callback implementation that renames each document part and saves it to a stream.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public SavedDocumentPartRename(string outputFolder, string baseFileName,
            DocumentSplitCriteria criteria)
        {
            _outputFolder = outputFolder;
            _baseFileName = baseFileName;
            _criteria = criteria;
            _count = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type based on the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Unknown"
            };

            // Build a unique file name for the part, preserving the original extension.
            string partFileName = $"{_baseFileName} part {++_count}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name and stream where Aspose.Words will write this part.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
