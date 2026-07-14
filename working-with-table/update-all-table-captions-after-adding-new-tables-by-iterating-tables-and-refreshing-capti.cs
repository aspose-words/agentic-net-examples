using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace TableCaptionUpdater
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add three tables, each preceded by a caption that uses a SEQ field.
            for (int i = 1; i <= 3; i++)
            {
                // Insert a SEQ field for tables. The field type is FieldSequence.
                Field field = builder.InsertField(FieldType.FieldSequence, true);
                // Set the sequence identifier to "Table" so that the field numbers tables.
                ((FieldSeq)field).SequenceIdentifier = "Table";

                // Write the caption text after the field.
                builder.Writeln($" - Sample Table {i}");

                // Build a simple 2x2 table.
                Table table = builder.StartTable();

                builder.InsertCell();
                builder.Write($"Row 1, Cell 1 (Table {i})");
                builder.InsertCell();
                builder.Write($"Row 1, Cell 2 (Table {i})");
                builder.EndRow();

                builder.InsertCell();
                builder.Write($"Row 2, Cell 1 (Table {i})");
                builder.InsertCell();
                builder.Write($"Row 2, Cell 2 (Table {i})");
                builder.EndRow();

                builder.EndTable();

                // Add a blank paragraph after each table for readability.
                builder.Writeln();
            }

            // Update all SEQ fields in the document so that caption numbers are correct.
            // This can be done globally, but we follow the requirement to iterate tables.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Find the nearest preceding paragraph (the caption).
                Node node = table.PreviousSibling;
                while (node != null && !(node is Paragraph))
                {
                    node = node.PreviousSibling;
                }

                if (node is Paragraph captionParagraph)
                {
                    // Update each SEQ (FieldSequence) field in the caption paragraph.
                    foreach (Field f in captionParagraph.Range.Fields)
                    {
                        if (f.Type == FieldType.FieldSequence)
                        {
                            f.Update();
                        }
                    }
                }
            }

            // Save the document.
            doc.Save("TableCaptionsUpdated.docx");
        }
    }
}
