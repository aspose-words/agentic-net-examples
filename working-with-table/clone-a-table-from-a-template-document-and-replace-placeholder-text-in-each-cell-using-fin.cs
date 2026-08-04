using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final output.
        string templatePath = "Template.docx";
        string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a table with placeholders.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Build a 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("{{Name}}");          // Placeholder for a person's name.
        builder.InsertCell();
        builder.Write("{{Age}}");           // Placeholder for age.

        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("{{City}}");          // Placeholder for city.
        builder.InsertCell();
        builder.Write("{{Country}}");       // Placeholder for country.

        builder.EndTable();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template, clone its table and import it into a new document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(templatePath);

        // Retrieve the first (and only) table from the template.
        Table sourceTable = (Table)sourceDoc.GetChildNodes(NodeType.Table, true)[0];

        // Clone the table. The cloned node still belongs to sourceDoc.
        Table clonedTable = (Table)sourceTable.Clone(true);

        // Create the destination document.
        Document destDoc = new Document();

        // Import the cloned table into the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(clonedTable, true);

        // Append the imported table to the destination document's body.
        destDoc.FirstSection.Body.AppendChild(importedTable);

        // -----------------------------------------------------------------
        // 3. Replace placeholders in each cell using FindReplaceOptions.
        // -----------------------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Perform replacements on the imported table's range.
        importedTable.Range.Replace("{{Name}}", "John Doe", replaceOptions);
        importedTable.Range.Replace("{{Age}}", "30", replaceOptions);
        importedTable.Range.Replace("{{City}}", "New York", replaceOptions);
        importedTable.Range.Replace("{{Country}}", "USA", replaceOptions);

        // -----------------------------------------------------------------
        // 4. Save the resulting document.
        // -----------------------------------------------------------------
        destDoc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
