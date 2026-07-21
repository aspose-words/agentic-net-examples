using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace UpdateTableCaptions
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add three tables, each preceded by a caption that uses a SEQ field.
            for (int i = 0; i < 3; i++)
            {
                // Insert a SEQ field for automatic table numbering.
                // The correct field type is FieldSequence and the sequence identifier is "Table".
                FieldSeq seqField = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
                seqField.SequenceIdentifier = "Table";

                // Write the rest of the caption text.
                builder.Writeln($": Sample Table {i + 1}");

                // Build a simple 2x2 table.
                Table table = builder.StartTable();
                for (int row = 0; row < 2; row++)
                {
                    for (int col = 0; col < 2; col++)
                    {
                        builder.InsertCell();
                        builder.Write($"R{row + 1}C{col + 1}");
                    }
                    builder.EndRow();
                }
                builder.EndTable();

                // Add a blank paragraph after each table for readability.
                builder.Writeln();
            }

            // Iterate through all tables (demonstration of traversal).
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                // Placeholder for any future table‑specific logic.
                _ = tbl; // Suppress unused variable warning.
            }

            // Refresh all SEQ fields (captions) to reflect the correct table numbers.
            doc.UpdateFields();

            // Save the resulting document.
            doc.Save("UpdatedTableCaptions.docx");
        }
    }
}
