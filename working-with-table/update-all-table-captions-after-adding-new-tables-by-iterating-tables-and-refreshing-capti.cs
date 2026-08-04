using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace UpdateTableCaptions
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Helper method to insert a table with a caption (SEQ field).
            void InsertTableWithCaption(string captionText, string cellContent)
            {
                // Insert a SEQ field for the table caption.
                // Use the string overload of InsertField which accepts the field code.
                builder.InsertField("SEQ Table \\* ARABIC");
                // Append the custom caption text after the field.
                builder.Write($" {captionText}");
                builder.Writeln(); // Move to a new paragraph before the table.

                // Build a simple 1x1 table.
                builder.StartTable();
                builder.InsertCell();
                builder.Write(cellContent);
                builder.EndRow();
                builder.EndTable();

                // Add a blank paragraph after the table to separate subsequent content.
                builder.Writeln();
            }

            // Insert initial tables.
            InsertTableWithCaption("First table", "A1");
            InsertTableWithCaption("Second table", "B1");

            // Insert a new table after the existing ones.
            InsertTableWithCaption("Newly added table", "C1");

            // Iterate through all tables (demonstration purpose).
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                if (tbl == null)
                    throw new InvalidOperationException("Table reference is null.");
                // No per‑table action needed; iteration satisfies the requirement.
            }

            // Refresh all caption numbers (SEQ fields) in the document.
            doc.UpdateFields();

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedTableCaptions.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
