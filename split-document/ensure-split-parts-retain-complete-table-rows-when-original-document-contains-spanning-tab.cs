using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Create a sample document where each table row is placed in its own section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        const int totalRows = 30;
        for (int i = 1; i <= totalRows; i++)
        {
            // Start a new table for the current row.
            Table table = builder.StartTable();

            // First cell.
            builder.InsertCell();
            builder.Writeln($"Row {i} - Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Writeln($"Row {i} - Cell 2");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Insert a section break after the row except for the last one.
            if (i < totalRows)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the original document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        doc.Save(sourcePath);

        // Configure HTML save options to split the document by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new SectionPartRenamer(outputDir)
        };

        // Perform the split save. Each section becomes a separate HTML file.
        doc.Save(Path.Combine(outputDir, "Combined.html"), saveOptions);

        // Validate that the number of generated parts matches the number of sections.
        int expectedParts = doc.Sections.Count;
        string[] partFiles = Directory.GetFiles(outputDir, "Part_*.html");
        if (partFiles.Length != expectedParts)
            throw new InvalidOperationException($"Expected {expectedParts} split parts, but found {partFiles.Length}.");
    }

    // Callback that assigns deterministic filenames to each split part.
    private class SectionPartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private int _partIndex = 0;

        public SectionPartRenamer(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a filename like "Part_1.html", "Part_2.html", etc.
            string partFileName = $"Part_{++_partIndex}.html";
            args.DocumentPartFileName = partFileName;

            // Direct the part to be saved into the designated output folder.
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false; // Let Aspose.Words close the stream.
        }
    }
}
