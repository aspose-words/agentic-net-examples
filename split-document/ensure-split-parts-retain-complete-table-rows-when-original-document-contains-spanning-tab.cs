using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitDocumentWithCompleteTableRows
{
    public static void Main()
    {
        // Output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // --------------------------------------------------------------------
        // Create a sample document that contains a table split across two sections.
        // The table is closed before the section break, then a new table is started
        // in the next section to simulate a logical table that spans sections.
        // --------------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First part of the table (section 1).
        Table table = builder.StartTable();
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");
            builder.EndRow();
        }
        builder.EndTable(); // Close the table before the break.

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second part of the same logical table (section 2).
        table = builder.StartTable();
        for (int i = 4; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");
            builder.EndRow();
        }
        builder.EndTable(); // Close the second part.

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // --------------------------------------------------------------------
        // Merge tables that are split across consecutive sections.
        // If a section ends with a table and the next section starts with a table,
        // move all rows from the next table into the previous one and remove the empty table.
        // --------------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count - 1; i++)
        {
            Section currentSection = sourceDoc.Sections[i];
            Section nextSection = sourceDoc.Sections[i + 1];

            // Get the last table of the current section (if any).
            Table currentTable = null;
            if (currentSection.Body.Tables.Count > 0)
                currentTable = currentSection.Body.Tables[currentSection.Body.Tables.Count - 1];

            // Get the first table of the next section (if any).
            Table nextTable = null;
            if (nextSection.Body.Tables.Count > 0)
                nextTable = nextSection.Body.Tables[0];

            // If both tables exist, treat them as parts of the same logical table.
            if (currentTable != null && nextTable != null)
            {
                // Move rows from the next table to the current table.
                foreach (Row row in nextTable.Rows)
                {
                    currentTable.Rows.Add(row.Clone(true));
                }

                // Remove the now‑empty table from the next section.
                nextTable.Remove();
            }
        }

        // --------------------------------------------------------------------
        // Split the document by sections. Each part will contain a complete
        // section (with any tables already merged), ensuring that no table row
        // is broken between parts.
        // --------------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document for the part.
            Document partDoc = new Document();

            // Import the section into the new document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            // Save the part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");
        }

        Console.WriteLine($"Document split into {sourceDoc.Sections.Count} parts. Files are located in: {outputDir}");
    }
}
