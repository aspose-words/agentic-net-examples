using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsNestedTableExtraction
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Create a source document that contains a table with a nested table.
            // -----------------------------------------------------------------
            Document srcDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(srcDoc);

            // Start the outer table.
            Table outerTable = builder.StartTable();

            // First row, first cell – this cell will contain a nested table.
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Writeln("Outer Table - Cell 1 (contains nested table)");

            // Create the nested table inside the current cell.
            Table nestedTable = builder.StartTable();

            // Nested table row 1.
            builder.InsertCell();
            builder.Writeln("Nested Row 1, Cell 1");
            builder.InsertCell();
            builder.Writeln("Nested Row 1, Cell 2");
            builder.EndRow();

            // Nested table row 2.
            builder.InsertCell();
            builder.Writeln("Nested Row 2, Cell 1");
            builder.InsertCell();
            builder.Writeln("Nested Row 2, Cell 2");
            builder.EndRow();

            // End the nested table.
            builder.EndTable();
            nestedTable.AllowAutoFit = true; // Apply after rows have been added.

            // Continue with the outer table – second cell of the first row.
            builder.InsertCell();
            builder.Writeln("Outer Table - Cell 2");
            builder.EndRow();

            // Second row of the outer table.
            builder.InsertCell();
            builder.Writeln("Outer Table - Row 2, Cell 1");
            builder.InsertCell();
            builder.Writeln("Outer Table - Row 2, Cell 2");
            builder.EndRow();

            // End the outer table.
            builder.EndTable();
            outerTable.AllowAutoFit = true; // Apply after rows have been added.

            // -----------------------------------------------------------------
            // Extract the outer table (including its nested tables) into a new document.
            // -----------------------------------------------------------------
            // Find the first table in the source document (the outer table we just created).
            Table extractedOuterTable = (Table)srcDoc.GetChild(NodeType.Table, 0, true);
            if (extractedOuterTable == null)
                throw new InvalidOperationException("No table found in the source document.");

            // Create a destination document.
            Document dstDoc = new Document();

            // Import the outer table (with nested tables) into the destination document.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(extractedOuterTable, true);

            // Append the imported table to the body of the destination document.
            dstDoc.FirstSection.Body.AppendChild(importedTable);

            // Save both documents for verification (optional).
            srcDoc.Save("SourceWithNestedTables.docx");
            dstDoc.Save("ExtractedSegment.docx");
        }
    }
}
