using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Content of Section 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage); // Start a new section.
        builder.Writeln("Content of Section 2.");

        // Set up HTML save options to split the document by sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SectionPartSaver(outputDir)
        };

        // Save the document; this will produce multiple HTML files.
        string mainFile = Path.Combine(outputDir, "Document.html");
        doc.Save(mainFile, saveOptions);

        // Validate that the expected split files were created.
        for (int i = 1; i <= doc.Sections.Count; i++)
        {
            string partPath = Path.Combine(outputDir, $"Section_{i}.html");
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split file not found: {partPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split completed successfully.");
    }

    // Callback that assigns custom filenames for each split part.
    private class SectionPartSaver : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private int _partIndex = 0;

        public SectionPartSaver(string outputDir)
        {
            _outputDir = outputDir;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;
            string fileName = $"Section_{_partIndex}.html";
            args.DocumentPartFileName = fileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, fileName), FileMode.Create);
        }
    }
}
