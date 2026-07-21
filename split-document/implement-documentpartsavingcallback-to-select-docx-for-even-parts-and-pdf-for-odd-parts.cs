using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with several sections to demonstrate splitting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add four sections, each will become a separate part when we split by SectionBreak.
        for (int i = 1; i <= 4; i++)
        {
            builder.Writeln($"Section {i}");
            if (i < 4)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Configure HtmlSaveOptions to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Assign the custom callback that decides the file format for each part.
        saveOptions.DocumentPartSavingCallback = new PartSavingCallback(outputDir, "SplitDocument");

        // Save the document; the callback will create separate files for each part.
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that the expected files were created.
        var createdFiles = Directory.GetFiles(outputDir)
                                    .Where(f => f.StartsWith(Path.Combine(outputDir, "SplitDocument_part")))
                                    .OrderBy(f => f)
                                    .ToArray();

        // Expect four parts: two DOCX (even indices 0,2) and two PDF (odd indices 1,3).
        if (createdFiles.Length != 4 ||
            !createdFiles[0].EndsWith(".docx") ||
            !createdFiles[1].EndsWith(".pdf") ||
            !createdFiles[2].EndsWith(".docx") ||
            !createdFiles[3].EndsWith(".pdf"))
        {
            throw new InvalidOperationException("The split parts were not created as expected.");
        }
    }

    // Callback implementation that selects DOCX for even parts and PDF for odd parts.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _baseName;
        private int _partIndex = 0;

        public PartSavingCallback(string outputDir, string baseName)
        {
            _outputDir = outputDir;
            _baseName = baseName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the desired extension based on part index parity.
            string extension = (_partIndex % 2 == 0) ? ".docx" : ".pdf";

            // Build a unique file name for the part.
            string partFileName = $"{_baseName}_part{_partIndex}{extension}";

            // Set the file name (without path) that Aspose.Words will use.
            args.DocumentPartFileName = partFileName;

            // Provide a stream that writes the part to the correct location.
            string fullPath = Path.Combine(_outputDir, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);

            // Ensure the stream is closed after the part is written.
            args.KeepDocumentPartStreamOpen = false;

            _partIndex++;
        }
    }
}
