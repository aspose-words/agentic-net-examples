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
        void InsertTableWithCaption(string captionText, string[,] data)
        {
            // Insert a paragraph for the caption.
            // Insert a SEQ field for table numbering.
            builder.Writeln(); // Ensure we are on a new paragraph.
            builder.InsertField("SEQ Table \\* ARABIC");
            // Append the caption text after the field.
            builder.Writeln($" {captionText}");
            builder.Writeln(); // Blank line before the table.

            // Build the table.
            builder.StartTable();
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    builder.InsertCell();
                    builder.Writeln(data[i, j]);
                }
                builder.EndRow();
            }
            builder.EndTable();
            builder.Writeln(); // Paragraph after the table.
        }

        // Insert initial tables.
        InsertTableWithCaption("Table Caption", new string[,] { { "A1", "B1" }, { "A2", "B2" } });
        InsertTableWithCaption("Table Caption", new string[,] { { "C1", "D1" }, { "C2", "D2" } });

        // Add a new table later in the document.
        InsertTableWithCaption("Table Caption", new string[,] { { "E1", "F1" }, { "E2", "F2" } });

        // Iterate all tables and refresh their caption numbers.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Find the preceding paragraph that should contain the caption.
            Node previousNode = table.PreviousSibling;
            while (previousNode != null && !(previousNode is Paragraph))
                previousNode = previousNode.PreviousSibling;

            if (previousNode is Paragraph captionParagraph)
            {
                // Update all fields (including SEQ fields) in the caption paragraph.
                foreach (Field field in captionParagraph.Range.Fields)
                {
                    field.Update();
                }
            }
        }

        // Ensure any remaining fields are up‑to‑date.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "UpdatedCaptions.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
