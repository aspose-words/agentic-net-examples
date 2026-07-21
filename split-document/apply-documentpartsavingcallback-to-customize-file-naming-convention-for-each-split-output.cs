using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 content.");

        // Set up HTML save options to split by section.
        string baseFileName = "SplitDocument.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new CustomPartNamingCallback(baseFileName, artifactsDir, DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; parts will be created via the callback.
        string mainOutputPath = Path.Combine(artifactsDir, baseFileName);
        doc.Save(mainOutputPath, saveOptions);

        // Verify that split parts were generated.
        var partFiles = Directory.GetFiles(artifactsDir, "SplitDocument_part*.html");
        if (partFiles.Length == 0)
            throw new Exception("No document parts were generated.");

        // Optional: display generated part file names.
        foreach (var file in partFiles)
            Console.WriteLine("Generated part: " + Path.GetFileName(file));
    }

    // Callback that renames each document part and saves it to a custom stream.
    private class CustomPartNamingCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly string _outputDir;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public CustomPartNamingCallback(string baseFileName, string outputDir, DocumentSplitCriteria criteria)
        {
            _baseFileName = Path.GetFileNameWithoutExtension(baseFileName);
            _outputDir = outputDir;
            _criteria = criteria;
            _count = 0;
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

            string newFileName = $"{_baseFileName}_part{++_count}_of_{partType}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = newFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, newFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
