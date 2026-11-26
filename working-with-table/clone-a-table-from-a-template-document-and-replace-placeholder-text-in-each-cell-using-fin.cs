using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a template document with a 2x2 table containing placeholder text.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("{{Placeholder}}");
        builder.InsertCell();
        builder.Write("{{Placeholder}}");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("{{Placeholder}}");
        builder.InsertCell();
        builder.Write("{{Placeholder}}");
        builder.EndRow();

        builder.EndTable();

        // Save the template document to a temporary file.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template document.
        var sourceDoc = new Document(templatePath);

        // Locate the first table in the source document.
        var sourceTable = sourceDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (sourceTable == null)
            throw new InvalidOperationException("No table found in the template document.");

        // Clone the table (deep clone).
        var clonedTable = (Table)sourceTable.Clone(true);

        // Replace placeholder text in each cell of the cloned table.
        var replaceOptions = new FindReplaceOptions();
        foreach (Row row in clonedTable.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.Range.Replace("{{Placeholder}}", "Replaced", replaceOptions);
            }
        }

        // Create a new document to receive the cloned table.
        var resultDoc = new Document();
        var resultBuilder = new DocumentBuilder(resultDoc);
        resultBuilder.Writeln("Table cloned and placeholders replaced:");

        // Import the cloned table into the result document before appending.
        var importer = new NodeImporter(sourceDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        var importedTable = (Table)importer.ImportNode(clonedTable, true);
        resultDoc.FirstSection.Body.AppendChild(importedTable);

        // Save the result document.
        const string resultPath = "Result.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file exists.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("Result document was not created.", resultPath);
    }
}
