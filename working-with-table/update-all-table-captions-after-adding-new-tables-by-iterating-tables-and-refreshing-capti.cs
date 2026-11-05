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

        // Helper method to insert a table with a caption.
        void InsertTableWithCaption(string captionText, int rows, int cols)
        {
            // Ensure the caption starts on a new paragraph.
            builder.Writeln();

            // Insert the caption: "Table {SEQ Table \\* ARABIC}: {captionText}"
            builder.Write("Table ");
            builder.InsertField("SEQ Table \\* ARABIC", "1");
            builder.Writeln($": {captionText}");

            // Build the table.
            builder.StartTable();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{r + 1}C{c + 1}");
                }
                builder.EndRow();
            }
            builder.EndTable();
        }

        // Insert initial tables with captions.
        InsertTableWithCaption("First table", 2, 2);
        InsertTableWithCaption("Second table", 3, 3);

        // Insert a new table after the initial ones.
        InsertTableWithCaption("Newly added table", 2, 3);

        // Refresh all table captions to reflect the new ordering.
        RefreshTableCaptions(doc);

        // Save the resulting document.
        string outputPath = "UpdatedCaptions.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }

    // Iterates through all tables and updates any preceding caption SEQ fields.
    private static void RefreshTableCaptions(Document doc)
    {
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Node node in tables)
        {
            Table table = (Table)node;

            // The caption paragraph is expected to be the immediate previous sibling.
            Paragraph captionParagraph = table.PreviousSibling as Paragraph;
            if (captionParagraph == null)
                continue;

            // Update each SEQ field that references Table captions.
            foreach (Field field in captionParagraph.Range.Fields)
            {
                // Check the field code for the SEQ Table identifier.
                if (field.GetFieldCode().Contains("SEQ Table"))
                {
                    field.Update();
                }
            }
        }
    }
}
