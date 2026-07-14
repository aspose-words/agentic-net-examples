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

        // Base file name for the main document and for generated parts.
        string baseFilePath = Path.Combine(outputDir, "SplitDocument.html");

        // Create a sample document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of the first section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of the second section.");

        // Configure HTML save options to split by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SavedDocumentPartRename(baseFilePath, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; the callback will rename each part.
        doc.Save(baseFilePath, saveOptions);

        // Optional: list the generated files.
        foreach (string file in Directory.GetFiles(outputDir, "SplitDocument*_part*.*"))
        {
            Console.WriteLine(file);
        }
    }

    // Callback that customizes the file name (and stream) for each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFilePath;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public SavedDocumentPartRename(string baseFilePath, DocumentSplitCriteria criteria)
        {
            _baseFilePath = baseFilePath;
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

            string extension = Path.GetExtension(args.DocumentPartFileName);
            string baseName = Path.GetFileNameWithoutExtension(_baseFilePath);
            string partFileName = $"{baseName}_part{++_count}_{partType}{extension}";

            // Set the new file name.
            args.DocumentPartFileName = partFileName;

            // Save the part to a custom stream in the same output directory.
            string directory = Path.GetDirectoryName(_baseFilePath);
            args.DocumentPartStream = new FileStream(Path.Combine(directory, partFileName), FileMode.Create);
        }
    }
}
