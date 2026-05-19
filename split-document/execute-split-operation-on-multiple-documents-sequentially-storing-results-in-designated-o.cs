using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Entry point.
    public static void Main()
    {
        // Root folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Input and output subfolders.
        string inputDir = Path.Combine(artifactsDir, "Input");
        string outputDir = Path.Combine(artifactsDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create two sample documents with different numbers of sections.
        string doc1Path = Path.Combine(inputDir, "Sample1.docx");
        string doc2Path = Path.Combine(inputDir, "Sample2.docx");
        CreateSampleDocument(doc1Path, 3, "Doc1");
        CreateSampleDocument(doc2Path, 4, "Doc2");

        // Split each document sequentially.
        SplitDocumentBySection(doc1Path, Path.Combine(outputDir, "Doc1"));
        SplitDocumentBySection(doc2Path, Path.Combine(outputDir, "Doc2"));
    }

    // Creates a simple document containing the specified number of sections.
    private static void CreateSampleDocument(string filePath, int sectionCount, string titlePrefix)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= sectionCount; i++)
        {
            builder.Writeln($"{titlePrefix} - Section {i}");
            // Insert a section break after each section except the last one.
            if (i < sectionCount)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        doc.Save(filePath);
    }

    // Splits the given document into separate HTML files, one per section.
    private static void SplitDocumentBySection(string sourcePath, string destinationFolder)
    {
        // Ensure the destination folder exists.
        Directory.CreateDirectory(destinationFolder);

        // Load the source document.
        Document doc = new Document(sourcePath);

        // Prepare HTML save options to split by section.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new PartSavingCallback(destinationFolder, Path.GetFileNameWithoutExtension(sourcePath))
        };

        // The main file name is required but its content will be ignored because we split into parts.
        string mainFilePath = Path.Combine(destinationFolder, $"{Path.GetFileNameWithoutExtension(sourcePath)}_Main.html");
        doc.Save(mainFilePath, options);

        // Validate that at least one part was created.
        var partFiles = Directory.GetFiles(destinationFolder, "*.html")
                                 .Where(f => !f.EndsWith("_Main.html", StringComparison.OrdinalIgnoreCase))
                                 .ToArray();

        if (partFiles.Length == 0)
            throw new InvalidOperationException($"No split parts were created for '{sourcePath}'.");
    }

    // Callback that controls how each document part is saved.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseName;
        private int _partIndex = 0;

        public PartSavingCallback(string outputFolder, string baseName)
        {
            _outputFolder = outputFolder;
            _baseName = baseName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a unique file name for the part.
            string partFileName = $"{_baseName}_Part{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (without path) and provide a stream that writes to the desired folder.
            args.DocumentPartFileName = partFileName;
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
