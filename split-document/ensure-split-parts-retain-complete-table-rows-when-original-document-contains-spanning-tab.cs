using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class SplitDocumentWithCompleteTableRows
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document containing a table that spans multiple pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with many rows to force pagination.
        Table table = builder.StartTable();
        for (int i = 1; i <= 30; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i}, Column 1");
            builder.InsertCell();
            builder.Writeln($"Row {i}, Column 2");
            builder.EndRow();
        }
        builder.EndTable();

        // Prevent rows from breaking across page boundaries.
        foreach (Row row in table.Rows)
            row.RowFormat.AllowBreakAcrossPages = false;

        // Save the document while splitting it into parts.
        // Use PageBreak criteria; rows will stay intact because of the setting above.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak
        };

        string baseFileName = "SplitDocument.html";
        saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(baseFileName, saveOptions.DocumentSplitCriteria);

        string mainOutputPath = Path.Combine(outputDir, baseFileName);
        doc.Save(mainOutputPath, saveOptions);

        // Validate that split parts were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument part *.html");
        if (splitFiles.Length == 0)
            throw new InvalidOperationException("No split document parts were generated.");

        // Display generated file names.
        foreach (string file in splitFiles)
            Console.WriteLine($"Generated: {Path.GetFileName(file)}");
    }

    // Callback to customize the filenames of each document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partCount;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = Path.GetFileNameWithoutExtension(baseFileName);
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

            string newFileName = $"{_baseFileName} part {++_partCount}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = newFileName;
            // Default stream handling is sufficient.
        }
    }
}
