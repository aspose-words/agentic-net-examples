using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace AsposeWordsTableCaptions
{
    public class Program
    {
        public static void Main()
        {
            // Path for the output document.
            const string outputPath = "UpdatedTableCaptions.docx";

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Helper method to insert a table with a preceding caption field.
            void InsertTableWithCaption(string captionText)
            {
                // Insert a paragraph that will hold the caption text.
                builder.Writeln(captionText);

                // Insert the SEQ field that generates the table number.
                // Use the string overload of InsertField to insert a field code.
                builder.InsertField(" SEQ Table \\* ARABIC ");

                builder.Writeln(); // Move to the next line after the field.

                // Build a simple 2‑row, 2‑column table.
                Table table = builder.StartTable();
                for (int r = 0; r < 2; r++)
                {
                    for (int c = 0; c < 2; c++)
                    {
                        builder.InsertCell();
                        builder.Write($"R{r + 1}C{c + 1}");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln(); // Add a blank line after the table.
            }

            // Insert a few tables with captions.
            InsertTableWithCaption("Table:");
            InsertTableWithCaption("Table:");
            InsertTableWithCaption("Table:");

            // Insert an additional table without a caption to simulate a later addition.
            builder.Writeln("Additional table without explicit caption:");
            Table extraTable = builder.StartTable();
            builder.InsertCell();
            builder.Write("Extra 1");
            builder.InsertCell();
            builder.Write("Extra 2");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln();

            // -----------------------------------------------------------------
            // Refresh all table captions.
            // -----------------------------------------------------------------
            // Iterate through every table in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // The caption field is usually placed in the paragraph immediately
                // preceding the table. Retrieve that paragraph.
                Paragraph captionParagraph = table.PreviousSibling as Paragraph;
                if (captionParagraph == null)
                    continue;

                // Find all SEQ fields inside the caption paragraph.
                NodeCollection fieldStarts = captionParagraph.GetChildNodes(NodeType.FieldStart, true);
                foreach (FieldStart fieldStart in fieldStarts)
                {
                    if (fieldStart.FieldType == FieldType.FieldSequence) // SEQ field type
                    {
                        // Retrieve the Field object from the FieldStart node and update it.
                        Field field = fieldStart.GetField();
                        field?.Update();
                    }
                }
            }

            // As a safety net, also update all fields in the document.
            doc.UpdateFields();

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }
    }
}
