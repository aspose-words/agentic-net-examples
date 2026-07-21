using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsPageOrientationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Section 1 – a normal table (portrait orientation).
            // -------------------------------------------------
            builder.Writeln("Section 1: Normal table (portrait).");
            builder.StartTable();
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"Cell {col + 1}");
            }
            builder.EndRow();
            builder.EndTable();

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // -------------------------------------------------
            // Section 2 – a wide table (will be set to landscape).
            // -------------------------------------------------
            builder.Writeln("Section 2: Wide table (landscape).");
            builder.StartTable();
            // Create a table with many columns to make it wide.
            for (int col = 0; col < 10; col++)
            {
                builder.InsertCell();
                builder.Write($"Wide {col + 1}");
            }
            builder.EndRow();
            builder.EndTable();

            // -------------------------------------------------
            // Detect sections that contain a "wide" table and set orientation to landscape.
            // For this example a table with more than 5 columns is considered wide.
            // -------------------------------------------------
            foreach (Section section in doc.Sections)
            {
                bool hasWideTable = false;

                // Get all tables in the current section.
                NodeCollection tables = section.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tables)
                {
                    // If the first row has more than 5 cells, treat it as a wide table.
                    if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > 5)
                    {
                        hasWideTable = true;
                        break;
                    }
                }

                // Apply landscape orientation if a wide table was found.
                if (hasWideTable)
                {
                    section.PageSetup.Orientation = Orientation.Landscape;
                }
            }

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "DocumentWithLandscapeSections.docx");
            doc.Save(outputPath);
        }
    }
}
