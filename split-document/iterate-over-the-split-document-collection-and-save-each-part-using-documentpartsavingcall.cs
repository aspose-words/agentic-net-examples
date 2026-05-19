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

        // Base name for the main file and split parts.
        string baseName = "SplitDocument";

        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Configure HTML save options to split by section.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(outputDir, baseName, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; the callback will be invoked for each part.
        string mainFilePath = Path.Combine(outputDir, baseName + ".html");
        doc.Save(mainFilePath, options);

        // Report the generated files.
        string[] generatedFiles = Directory.GetFiles(outputDir, $"{baseName}*");
        Console.WriteLine($"Generated {generatedFiles.Length} files:");
        foreach (string file in generatedFiles)
        {
            Console.WriteLine(Path.GetFileName(file));
        }
    }

    // Callback that customizes the file name and stream for each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public SavedDocumentPartRename(string outputDir, string baseName, DocumentSplitCriteria criteria)
        {
            _outputDir = outputDir;
            _baseName = baseName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            string partFileName = $"{_baseName}_part{++_count}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name that Aspose.Words will use.
            args.DocumentPartFileName = partFileName;

            // Provide a custom stream to write the part.
            string fullPath = Path.Combine(_outputDir, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }
}
