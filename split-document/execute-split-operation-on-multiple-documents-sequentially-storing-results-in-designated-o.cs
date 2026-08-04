using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Creates a sample DOCX file with the specified number of sections.
    private static void CreateSampleDocument(string filePath, int sectionCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= sectionCount; i++)
        {
            // Insert a heading for each section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Section {i}");

            // Insert some body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of section {i}.");

            // Add a section break after each section except the last one.
            if (i < sectionCount)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Ensure the directory exists and save the document.
        Directory.CreateDirectory(Path.GetDirectoryName(filePath));
        doc.Save(filePath);
    }

    // Callback that saves each split part into a designated folder.
    private class DocumentPartSaver : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private int _partIndex = 0;

        public DocumentPartSaver(string outputFolder)
        {
            _outputFolder = outputFolder;
            Directory.CreateDirectory(_outputFolder);
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;
            string partFileName = $"Part_{_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";
            string fullPath = Path.Combine(_outputFolder, partFileName);

            // Set the stream where Aspose.Words will write this part.
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            // The file name property is not used when a stream is supplied, but set it for completeness.
            args.DocumentPartFileName = partFileName;
        }
    }

    public static void Main()
    {
        // Base directories for input documents and split output.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(inputDir);

        // Create sample source documents.
        string docPath1 = Path.Combine(inputDir, "Sample1.docx");
        string docPath2 = Path.Combine(inputDir, "Sample2.docx");
        CreateSampleDocument(docPath1, 3); // 3 sections
        CreateSampleDocument(docPath2, 2); // 2 sections

        // List of documents to process.
        List<string> sourceDocs = new List<string> { docPath1, docPath2 };

        foreach (string sourcePath in sourceDocs)
        {
            // Load the source document.
            Document doc = new Document(sourcePath);
            string docName = Path.GetFileNameWithoutExtension(sourcePath);

            // Folder where split parts for this document will be stored.
            string docOutputFolder = Path.Combine(baseDir, docName);
            Directory.CreateDirectory(docOutputFolder);

            // Configure HTML save options to split by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new DocumentPartSaver(docOutputFolder)
            };

            // Main HTML file (required by the Save method). Parts are written via the callback.
            string mainFilePath = Path.Combine(docOutputFolder, $"{docName}.html");
            doc.Save(mainFilePath, saveOptions);

            // Verify that at least one split part was created.
            string[] partFiles = Directory.GetFiles(docOutputFolder, "Part_*.html");
            if (partFiles.Length == 0)
                throw new Exception($"No split parts were generated for document '{docName}'.");
        }

        // Execution completed without interactive prompts.
    }
}
