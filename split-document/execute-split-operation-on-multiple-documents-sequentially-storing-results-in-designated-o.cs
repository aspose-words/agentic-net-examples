using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directories for source documents and split results.
        string baseDir = Directory.GetCurrentDirectory();
        string sourceDir = Path.Combine(baseDir, "SourceDocs");
        string outputDir = Path.Combine(baseDir, "SplitResults");

        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // Create two sample documents with three sections each.
        string[] docNames = { "Sample1.docx", "Sample2.docx" };
        foreach (string docName in docNames)
        {
            string docPath = Path.Combine(sourceDir, docName);
            CreateSampleDocument(docPath, new[]
            {
                "Section A of " + Path.GetFileNameWithoutExtension(docName),
                "Section B of " + Path.GetFileNameWithoutExtension(docName),
                "Section C of " + Path.GetFileNameWithoutExtension(docName)
            });
        }

        // Process each document sequentially and split by section.
        foreach (string docName in docNames)
        {
            string sourcePath = Path.Combine(sourceDir, docName);
            Document doc = new Document(sourcePath);

            // Prepare output folder for this document.
            string docOutputFolder = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docName));
            Directory.CreateDirectory(docOutputFolder);

            // Base file name for the first part (Aspose will generate additional parts).
            string baseFileName = Path.Combine(docOutputFolder, Path.GetFileNameWithoutExtension(docName) + ".html");

            // Configure split options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };
            saveOptions.DocumentPartSavingCallback = new PartSavingCallback(baseFileName, saveOptions.DocumentSplitCriteria);

            // Save the document – this will invoke the callback for each part.
            doc.Save(baseFileName, saveOptions);

            // Validate that at least one split part was created.
            string[] createdParts = Directory.GetFiles(docOutputFolder, "*.html");
            if (createdParts.Length == 0)
                throw new InvalidOperationException($"No split files were created for '{docName}'.");
        }
    }

    // Creates a simple document with the supplied section texts.
    private static void CreateSampleDocument(string filePath, string[] sectionTexts)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < sectionTexts.Length; i++)
        {
            builder.Writeln(sectionTexts[i]);

            // Insert a section break after each section except the last.
            if (i < sectionTexts.Length - 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        doc.Save(filePath);
    }

    // Callback that assigns custom file names and streams for each split part.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseFilePath;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public PartSavingCallback(string baseFilePath, DocumentSplitCriteria criteria)
        {
            _baseFilePath = baseFilePath;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string partType = _criteria switch
            {
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            _partIndex++;
            string baseNameWithoutExt = Path.GetFileNameWithoutExtension(_baseFilePath);
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string partFileName = $"{baseNameWithoutExt} part {_partIndex}, of type {partType}{extension}";

            // Set the file name (without path) and provide a stream that writes to the correct folder.
            args.DocumentPartFileName = partFileName;
            string folder = Path.GetDirectoryName(_baseFilePath);
            args.DocumentPartStream = new FileStream(Path.Combine(folder, partFileName), FileMode.Create);
        }
    }
}
