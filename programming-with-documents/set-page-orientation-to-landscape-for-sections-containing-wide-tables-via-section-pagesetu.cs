using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a section with a narrow table (will stay portrait).
        builder.Writeln("Section 1 - Narrow Table");
        Table narrowTable = builder.StartTable();
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Cell {i + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add a section with a wide table (more columns, likely to need landscape).
        builder.Writeln("Section 2 - Wide Table");
        Table wideTable = builder.StartTable();
        for (int i = 0; i < 10; i++) // 10 columns makes the table wide.
        {
            builder.InsertCell();
            builder.Write($"Col {i + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Iterate through each section and set orientation to landscape
        // if the section contains at least one table with more than 5 columns.
        foreach (Section section in doc.Sections)
        {
            bool hasWideTable = false;
            foreach (Node node in section.Body.GetChildNodes(NodeType.Table, true))
            {
                Table tbl = (Table)node;
                if (tbl.Rows.Count > 0 && tbl.Rows[0].Cells.Count > 5)
                {
                    hasWideTable = true;
                    break;
                }
            }

            if (hasWideTable)
            {
                section.PageSetup.Orientation = Orientation.Landscape;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was saved and can be reopened.
        Document reopened = new Document(outputPath);
        // No further action needed; the program ends here.
    }
}
