using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1 – normal orientation with a small table.
        builder.Writeln("Section 1 - normal table");
        builder.StartTable();
        for (int i = 0; i < 2; i++)
        {
            builder.InsertCell();
            builder.Write($"R1C{i + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 – contains a wide table (many columns).
        builder.Writeln("Section 2 - wide table");
        builder.StartTable();
        int columnCount = 10; // Wide table threshold.
        // Header row.
        for (int i = 0; i < columnCount; i++)
        {
            builder.InsertCell();
            builder.Write($"Header {i + 1}");
        }
        builder.EndRow();
        // Data row.
        for (int i = 0; i < columnCount; i++)
        {
            builder.InsertCell();
            builder.Write($"Data {i + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // After the document is built, set orientation to landscape for any section that contains a wide table.
        foreach (Section section in doc.Sections)
        {
            bool containsWideTable = false;

            // Look for Table nodes inside the current section.
            NodeCollection tables = section.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Heuristic: treat tables with more than 5 columns as "wide".
                if (table.Rows.Count > 0 && table.FirstRow.Cells.Count > 5)
                {
                    containsWideTable = true;
                    break;
                }
            }

            if (containsWideTable)
            {
                // Change the page orientation for this section.
                section.PageSetup.Orientation = Orientation.Landscape;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "WideTableLandscape.docx");
        doc.Save(outputPath);
    }
}
