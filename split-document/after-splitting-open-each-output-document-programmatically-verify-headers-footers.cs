using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder where the split parts will be written.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with two sections, each having a header and a footer.
        Document srcDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(srcDoc);

        // First section
        builder.Writeln("Section 1 - Content");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer 1");
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section
        builder.Writeln("Section 2 - Content");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer 2");

        // Base name for the output HTML files (the main file name without extension).
        const string outputBaseName = "SplitDocument.html";

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Attach a custom callback that will name each part and keep track of the generated files.
        var partCallback = new PartSavingCallback(outputBaseName, artifactsDir);
        saveOptions.DocumentPartSavingCallback = partCallback;

        // Save the document – this will invoke the callback for each part.
        srcDoc.Save(Path.Combine(artifactsDir, outputBaseName), saveOptions);

        // After saving, verify that each generated part contains the expected headers/footers.
        foreach (string partPath in partCallback.GeneratedPartPaths)
        {
            // Load the part as a separate document.
            Document partDoc = new Document(partPath);

            // Verify that the first section has at least one header and one footer.
            bool hasHeader = partDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary] != null;
            bool hasFooter = partDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary] != null;

            Debug.Assert(hasHeader, $"Header missing in part: {partPath}");
            Debug.Assert(hasFooter, $"Footer missing in part: {partPath}");
        }

        Console.WriteLine("Document split and verification completed successfully.");
    }

    // Callback that assigns a unique filename to each document part and records the full path.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly string _outputFolder;
        private int _counter;
        public List<string> GeneratedPartPaths { get; } = new List<string>();

        public PartSavingCallback(string baseName, string outputFolder)
        {
            _baseName = baseName;
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Create a unique filename for the part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseName)}_part{++_counter}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename and stream where Aspose.Words will write the part.
            args.DocumentPartFileName = partFileName;
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);

            // Record the full path for later verification.
            GeneratedPartPaths.Add(fullPath);
        }
    }
}
