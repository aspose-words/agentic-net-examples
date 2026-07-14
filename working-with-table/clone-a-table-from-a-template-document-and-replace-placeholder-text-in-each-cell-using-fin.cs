using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        string outputPath = Path.Combine(artifactsDir, "ClonedTable.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a table with placeholders.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

        // Build a simple 2‑cell table.
        tmplBuilder.StartTable();
        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Name}}");          // Placeholder for a name.
        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Age}}");           // Placeholder for an age.
        tmplBuilder.EndRow();
        tmplBuilder.EndTable();

        // Save the template so it can be loaded later.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document and locate the table to clone.
        // -----------------------------------------------------------------
        Document srcDoc = new Document(templatePath);
        Table srcTable = srcDoc.GetChildNodes(NodeType.Table, true)[0] as Table;

        // -----------------------------------------------------------------
        // 3. Clone the table (deep clone) so we can modify it independently.
        // -----------------------------------------------------------------
        Table clonedTable = srcTable.Clone(true) as Table;

        // -----------------------------------------------------------------
        // 4. Replace placeholder text inside each cell using FindReplaceOptions.
        // -----------------------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Replace the placeholders with actual values.
        clonedTable.Range.Replace("{{Name}}", "John Doe", replaceOptions);
        clonedTable.Range.Replace("{{Age}}", "30", replaceOptions);

        // -----------------------------------------------------------------
        // 5. Insert the cloned and modified table into a new document.
        // -----------------------------------------------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Cloned table with replaced values:");

        // Import the cloned table into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(clonedTable, true);
        dstDoc.FirstSection.Body.AppendChild(importedTable);

        // -----------------------------------------------------------------
        // 6. Save the resulting document.
        // -----------------------------------------------------------------
        dstDoc.Save(outputPath);
    }
}
