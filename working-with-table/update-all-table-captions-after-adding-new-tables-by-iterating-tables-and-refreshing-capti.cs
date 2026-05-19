using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert a table with a caption.
        void InsertTableWithCaption(string captionText, int rows, int cols)
        {
            // Insert caption paragraph with a SEQ field for tables.
            builder.Writeln();
            Paragraph captionPara = builder.CurrentParagraph;
            Field seqField = builder.InsertField("SEQ Table \\* ARABIC", "1");
            captionPara.AppendChild(new Run(doc, " " + captionText));
            // Build the table.
            builder.StartTable();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    builder.InsertCell();
                    builder.Write($"R{r + 1}C{c + 1}");
                }
                builder.EndRow();
            }
            builder.EndTable();
        }

        // Insert initial tables.
        InsertTableWithCaption("First table caption.", 2, 3);
        InsertTableWithCaption("Second table caption.", 3, 2);

        // Add a new table later in the document.
        InsertTableWithCaption("Third table caption added later.", 2, 2);

        // Iterate all tables to demonstrate traversal.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        Console.WriteLine($"Total tables before updating captions: {tables.Count}");

        // Refresh caption numbers by updating fields.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedCaptions.docx");
        doc.Save(outputPath);

        // Verify the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
