using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Callback that controls how each split part is saved.
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
            // Generate a deterministic file name for the part.
            string partFileName = $"{_baseName}_Part{++_partIndex}.html";

            // Set the file name (without path) – Aspose will use the directory of the main file.
            args.DocumentPartFileName = partFileName;

            // Provide a stream that writes directly into the desired folder.
            string fullPath = Path.Combine(_outputDir, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    public static void Main()
    {
        // Base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(baseDir);

        // Create and process two sample documents.
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            // -----------------------------------------------------------------
            // 1. Create a sample source document with two sections.
            // -----------------------------------------------------------------
            string sourcePath = Path.Combine(baseDir, $"SourceDoc{docIndex}.docx");
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln($"Document {docIndex} - Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln($"Document {docIndex} - Section 2");

            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Load the document and split it by section using HTML save options.
            // -----------------------------------------------------------------
            Document docToSplit = new Document(sourcePath);

            // Folder where split parts will be stored.
            string partsDir = Path.Combine(baseDir, $"Doc{docIndex}_Parts");
            Directory.CreateDirectory(partsDir);

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new PartSavingCallback(partsDir, $"Doc{docIndex}")
            };

            // The main file name (required by Aspose to resolve relative paths).
            string mainFilePath = Path.Combine(partsDir, $"Doc{docIndex}.html");
            docToSplit.Save(mainFilePath, saveOptions);

            // -----------------------------------------------------------------
            // 3. Validate that split files were created.
            // -----------------------------------------------------------------
            string[] htmlFiles = Directory.GetFiles(partsDir, "*.html");
            // Expect at least two files: the main file and one additional part.
            if (htmlFiles.Length < 2)
                throw new InvalidOperationException($"Expected split files for document {docIndex} were not created.");

            Console.WriteLine($"Document {docIndex} split into {htmlFiles.Length} parts:");
            foreach (string file in htmlFiles)
                Console.WriteLine($"  {Path.GetFileName(file)}");
        }

        Console.WriteLine("All documents processed successfully.");
    }
}
