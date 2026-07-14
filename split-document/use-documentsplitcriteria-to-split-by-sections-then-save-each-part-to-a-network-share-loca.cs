using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Simulate a network share location (use a local folder for the demo).
        string networkShareFolder = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkShareFolder);

        // Base file name for the main HTML document.
        string baseFileName = "SplitDocument.html";

        // Configure HTML save options to split by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new NetworkSharePartSavingCallback(networkShareFolder, baseFileName, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; parts will be written to the network share folder.
        string mainOutputPath = Path.Combine(networkShareFolder, baseFileName);
        doc.Save(mainOutputPath, saveOptions);

        // Verify that at least one split part was created.
        string[] partFiles = Directory.GetFiles(networkShareFolder, $"{Path.GetFileNameWithoutExtension(baseFileName)} part*{Path.GetExtension(baseFileName)}");
        if (partFiles.Length == 0)
            throw new InvalidOperationException("No document parts were saved.");

        // Optional: output the list of created files (not required for the task).
        foreach (string file in partFiles)
            Console.WriteLine($"Created part: {file}");
    }

    // Callback that redirects each document part to the specified network share folder.
    private class NetworkSharePartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _folder;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex;

        public NetworkSharePartSavingCallback(string folder, string baseFileName, DocumentSplitCriteria criteria)
        {
            _folder = folder;
            _baseFileName = baseFileName;
            _criteria = criteria;
            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name (optional, for naming).
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)} part {++_partIndex} of {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (without path) and provide a stream pointing to the network share folder.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_folder, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
