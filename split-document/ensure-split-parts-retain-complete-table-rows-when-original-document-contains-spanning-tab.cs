using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    // Callback to give each split part a deterministic file name.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private int _partIndex = 0;

        public PartRenamer(string baseName)
        {
            _baseName = baseName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a name like "Split_Part_1.html", "Split_Part_2.html", …
            string partFileName = $"{_baseName}_Part_{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = partFileName;
        }
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document where each table row lives in its own section.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        Document srcDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(srcDoc);

        builder.Writeln("Table with rows split by section breaks (each row stays intact).");
        builder.Writeln();

        // Create a separate table (with a single row) for each section.
        for (int i = 1; i <= 5; i++)
        {
            // Start a new table.
            Table table = builder.StartTable();

            // First cell.
            builder.InsertCell();
            builder.Write($"Row {i} - Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Write($"Row {i} - Cell 2");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Insert a section break so that the next row lives in its own section.
            // No break after the last row.
            if (i < 5)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document.
        srcDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by section while preserving complete table rows.
        // -----------------------------------------------------------------
        Document docToSplit = new Document(sourcePath);

        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        string baseOutputName = "Split";
        saveOptions.DocumentPartSavingCallback = new PartRenamer(baseOutputName);

        string mainOutputPath = Path.Combine(Directory.GetCurrentDirectory(), $"{baseOutputName}.html");
        docToSplit.Save(mainOutputPath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Verify that split parts were created.
        // -----------------------------------------------------------------
        Console.WriteLine("Main HTML file: " + mainOutputPath);
        for (int i = 1; i <= docToSplit.Sections.Count; i++)
        {
            string partPath = Path.Combine(Directory.GetCurrentDirectory(),
                $"{baseOutputName}_Part_{i}.html");
            if (File.Exists(partPath))
                Console.WriteLine($"Part {i} created: {partPath}");
            else
                throw new FileNotFoundException($"Expected split part not found: {partPath}");
        }

        // Optional cleanup (commented out).
        // File.Delete(sourcePath);
        // File.Delete(mainOutputPath);
        // for (int i = 1; i <= docToSplit.Sections.Count; i++)
        //     File.Delete(Path.Combine(Directory.GetCurrentDirectory(),
        //         $"{baseOutputName}_Part_{i}.html"));
    }
}
