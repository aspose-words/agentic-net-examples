using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class TableConditionalColumnExample
{
    static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

        // Apply conditional formatting to the last column: make text bold.
        tableStyle.ConditionalStyles.LastColumn.Font.Bold = true;

        // Assign the style to the table.
        table.Style = tableStyle;

        // Enable the LastColumn conditional style on this table.
        table.StyleOptions |= TableStyleOptions.LastColumn;

        // Save the document as plain text while preserving table layout.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };
        string txtPath = Path.Combine(artifactsDir, "TableWithConditionalColumn.txt");
        doc.Save(txtPath, txtOptions);

        // Also save as DOCX for reference.
        string docxPath = Path.Combine(artifactsDir, "TableWithConditionalColumn.docx");
        doc.Save(docxPath);
    }
}
