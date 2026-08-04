using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Simulated network share folder (replace with actual UNC path if needed)
        string networkSharePath = Path.Combine(Environment.CurrentDirectory, "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Create a sample document with three sections
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Configure HTML save options to split by section breaks
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(networkSharePath, "SplitDocument.html")
        };

        // Save the document; each section will be saved as a separate HTML file in the network share folder
        string mainFilePath = Path.Combine(networkSharePath, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that the expected number of split files were created
        string[] splitFiles = Directory.GetFiles(networkSharePath, "SplitDocument_part*.html");
        if (splitFiles.Length != doc.Sections.Count)
            throw new InvalidOperationException($"Expected {doc.Sections.Count} split files, but found {splitFiles.Length}.");

        // Program ends automatically
    }

    // Callback to control how each document part is saved
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string outputFolder, string baseFileName)
        {
            _outputFolder = outputFolder;
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a unique filename for each part
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename (optional) and the stream where the part will be written
            args.DocumentPartFileName = partFileName;
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
