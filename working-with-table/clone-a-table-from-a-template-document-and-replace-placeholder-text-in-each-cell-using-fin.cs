using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the template and result documents.
        const string templatePath = "Template.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a template document that contains a table with
        // placeholder text in each cell.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Build a simple 2x2 table with placeholder text "{{Placeholder}}".
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("{{Placeholder}}");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("{{Placeholder}}");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("{{Placeholder}}");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("{{Placeholder}}");
        builder.EndRow();

        builder.EndTable();

        // Save the template document.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template document and clone its first table.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Table originalTable = (Table)loadedTemplate.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("No table found in the template document.");

        // Deep clone the table (including its contents).
        Table clonedTable = (Table)originalTable.Clone(true);

        // -----------------------------------------------------------------
        // Step 3: Create a destination document and import the cloned table.
        // -----------------------------------------------------------------
        Document destDoc = new Document(); // Empty document with a single section.

        // Import the cloned table into the destination document.
        NodeImporter importer = new NodeImporter(loadedTemplate, destDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTableNode = importer.ImportNode(clonedTable, true);
        Table importedTable = (Table)importedTableNode;

        // Append the imported table to the body of the destination document.
        destDoc.FirstSection.Body.AppendChild(importedTable);

        // -----------------------------------------------------------------
        // Step 4: Replace placeholder text in each cell using FindReplaceOptions.
        // -----------------------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        // Example: make the replacement case‑insensitive.
        replaceOptions.MatchCase = false;

        foreach (Row row in importedTable.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Replace the placeholder with actual content.
                cell.Range.Replace("{{Placeholder}}", "Replaced Text", replaceOptions);
            }
        }

        // -----------------------------------------------------------------
        // Step 5: Save the resulting document.
        // -----------------------------------------------------------------
        destDoc.Save(resultPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException($"Failed to create the result document at '{resultPath}'.");
    }
}
