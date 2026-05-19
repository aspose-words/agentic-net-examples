using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class SplitDocumentWithTableRows
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document containing a table that spans a section break.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert the first part of the table (rows 1‑10).
        builder.StartTable();
        for (int i = 1; i <= 20; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} - Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i} - Cell 2");
            builder.EndRow();

            // After the 10th row close the current table, insert a section break,
            // and start a new table so that the logical table spans two sections.
            if (i == 10)
            {
                builder.EndTable(); // End the first table.
                builder.InsertBreak(BreakType.SectionBreakNewPage); // Section break.
                builder.StartTable(); // Start the second part of the logical table.
            }
        }
        builder.EndTable(); // Close the final table.

        // Prevent rows from breaking across pages.
        foreach (Table tbl in sourceDoc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row row in tbl.Rows)
                row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Configure HTML save options to split by section.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Use a custom callback to give each split part a deterministic file name.
        string baseFileName = "SplitDocument";
        saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(baseFileName, saveOptions.DocumentSplitCriteria, artifactsDir);

        // Save the document; Aspose.Words will invoke the callback for each part.
        string mainOutputPath = Path.Combine(artifactsDir, $"{baseFileName}.html");
        sourceDoc.Save(mainOutputPath, saveOptions);

        // Gather all generated HTML parts.
        string[] partFiles = Directory.GetFiles(artifactsDir, $"{baseFileName}_Part*.html")
                                      .OrderBy(f => f)
                                      .ToArray();

        // Include the main part if it was also renamed by the callback.
        string mainPart = Path.Combine(artifactsDir, $"{baseFileName}_Part1.html");
        if (File.Exists(mainPart) && !partFiles.Contains(mainPart))
        {
            partFiles = new[] { mainPart }.Concat(partFiles).ToArray();
        }

        if (partFiles.Length == 0)
            throw new Exception("No split parts were generated.");

        // Validate that each part contains whole table rows only.
        int totalRows = 0;
        foreach (string partPath in partFiles)
        {
            Document partDoc = new Document(partPath);
            var tables = partDoc.GetChildNodes(NodeType.Table, true);
            int rowsInPart = 0;
            foreach (Table tbl in tables)
                rowsInPart += tbl.Rows.Count;

            if (rowsInPart == 0)
                throw new Exception($"Part '{Path.GetFileName(partPath)}' contains no table rows.");

            totalRows += rowsInPart;
        }

        // Verify that the total rows across all parts equal the original row count (20).
        if (totalRows != 20)
            throw new Exception($"Row count mismatch after splitting. Expected 20, found {totalRows}.");

        Console.WriteLine("Document split successfully. All parts retain complete table rows.");
    }

    // Callback that assigns deterministic file names to each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private readonly string _outputFolder;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseName, DocumentSplitCriteria criteria, string outputFolder)
        {
            _baseName = baseName;
            _criteria = criteria;
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a name like "SplitDocument_Part1.html", "SplitDocument_Part2.html", etc.
            string partFileName = $"{_baseName}_Part{++_partIndex}.html";

            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
