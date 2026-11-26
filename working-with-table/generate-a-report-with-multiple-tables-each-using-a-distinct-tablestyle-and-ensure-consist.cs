using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableReportExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- First Table ----------
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Data rows.
            for (int i = 1; i <= 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i} Col 1");
                builder.InsertCell();
                builder.Write($"Row {i} Col 2");
                builder.EndRow();
            }

            builder.EndTable();

            // Apply a distinct style to the first table.
            Table firstTable = (Table)doc.GetChildNodes(NodeType.Table, true)[0];
            firstTable.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

            // Consistent spacing after the first table.
            builder.ParagraphFormat.SpaceAfter = 12; // 12 points spacing.
            builder.Writeln(); // Insert an empty paragraph with the spacing.
            builder.ParagraphFormat.SpaceAfter = 0; // Reset spacing for following content.

            // ---------- Second Table ----------
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Column A");
            builder.InsertCell();
            builder.Write("Column B");
            builder.InsertCell();
            builder.Write("Column C");
            builder.EndRow();

            // Data rows.
            for (int i = 1; i <= 2; i++)
            {
                builder.InsertCell();
                builder.Write($"Item {i} A");
                builder.InsertCell();
                builder.Write($"Item {i} B");
                builder.InsertCell();
                builder.Write($"Item {i} C");
                builder.EndRow();
            }

            builder.EndTable();

            // Apply a different style to the second table.
            Table secondTable = (Table)doc.GetChildNodes(NodeType.Table, true)[1];
            secondTable.StyleIdentifier = StyleIdentifier.LightListAccent2;

            // Save the document.
            string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
