using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- First section: normal table (portrait) ----------
            builder.Writeln("Section 1 – normal table (portrait).");
            builder.StartTable();
            // Two columns.
            for (int i = 0; i < 2; i++)
            {
                builder.InsertCell();
                builder.Write($"Cell {i + 1}");
            }
            builder.EndRow();
            builder.EndTable();

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ---------- Second section: wide table ----------
            builder.Writeln("Section 2 – wide table (should become landscape).");
            builder.StartTable();
            // Five columns – considered wide.
            for (int i = 0; i < 5; i++)
            {
                builder.InsertCell();
                builder.Write($"Cell {i + 1}");
            }
            builder.EndRow();
            builder.EndTable();

            // Iterate through all sections and set orientation to Landscape
            // for any section that contains a table with more than 3 columns.
            foreach (Section section in doc.Sections)
            {
                bool hasWideTable = false;

                // Search for Table nodes inside the section.
                NodeCollection tables = section.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tables)
                {
                    // If the table has more than 3 columns, treat it as wide.
                    if (table.Rows.Count > 0 && table.FirstRow.Cells.Count > 3)
                    {
                        hasWideTable = true;
                        break;
                    }
                }

                if (hasWideTable)
                {
                    // Set the page orientation for this section to Landscape.
                    section.PageSetup.Orientation = Orientation.Landscape;
                }
            }

            // Save the document to the local file system.
            string outputPath = "WideTableOrientation.docx";
            doc.Save(outputPath);
        }
    }
}
