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
        // Input document to be split.
        const string inputFile = @"MyDir\Input.docx";

        // Base name for the output HTML files (the main file name without extension).
        const string outputBaseName = "SplitDocument.html";

        // Folder where the split parts will be written.
        const string artifactsDir = @"ArtifactsDir\";

        // Ensure the output folder exists.
        if (!Directory.Exists(artifactsDir))
            Directory.CreateDirectory(artifactsDir);

        // Load the source document.
        Document srcDoc = new Document(inputFile);

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

            // Example verification: ensure that the first section has at least one header and one footer.
            bool hasHeader = partDoc.FirstSection.HeadersFooters.Count > 0;
            bool hasFooter = partDoc.FirstSection.HeadersFooters.Count > 0; // same collection holds both

            Debug.Assert(hasHeader, $"Header missing in part: {partPath}");
            Debug.Assert(hasFooter, $"Footer missing in part: {partPath}");

            // Additional checks can be performed here, e.g., specific header types:
            // var header = partDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            // Debug.Assert(header != null && !string.IsNullOrWhiteSpace(header.GetText().Trim()));
        }
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
