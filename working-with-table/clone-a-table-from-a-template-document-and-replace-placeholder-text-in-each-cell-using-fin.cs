using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final output.
        string templatePath = "Template.docx";
        string resultPath = "Result.docx";

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
        tmplBuilder.Write("{{Address}}");       // Placeholder for an address.
        tmplBuilder.EndRow();
        tmplBuilder.EndTable();

        // Save the template so that it mimics an external source file.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document (simulating a real‑world scenario).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Locate the first table in the template.
        Table templateTable = (Table)loadedTemplate.GetChild(NodeType.Table, 0, true);
        if (templateTable == null)
            throw new InvalidOperationException("Template does not contain a table.");

        // -----------------------------------------------------------------
        // 3. Create the destination document and import (clone) the table.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();

        // Use NodeImporter to copy the table from the template into the result document.
        NodeImporter importer = new NodeImporter(loadedTemplate, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(templateTable, true);

        // Append the imported table to the body of the result document.
        resultDoc.FirstSection.Body.AppendChild(importedTable);

        // -----------------------------------------------------------------
        // 4. Replace placeholder text inside each cell using FindReplaceOptions.
        // -----------------------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Replace the placeholders throughout the whole table.
        importedTable.Range.Replace("{{Name}}", "John Doe", replaceOptions);
        importedTable.Range.Replace("{{Address}}", "123 Main St, Anytown", replaceOptions);

        // -----------------------------------------------------------------
        // 5. Save the final document.
        // -----------------------------------------------------------------
        resultDoc.Save(resultPath);

        // Simple validation that the file was created.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("The result document was not saved.", resultPath);
    }
}
