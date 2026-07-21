using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare folders and file paths.
        // -----------------------------------------------------------------
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        string outputPath   = Path.Combine(artifactsDir, "Result.docx");

        // -----------------------------------------------------------------
        // 2. Create a template document that contains a 2x2 table with
        //    placeholder text in each cell.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

        // Build the table.
        Table tmplTable = tmplBuilder.StartTable();
        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Name}}");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Age}}");
        tmplBuilder.EndRow();

        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{City}}");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Country}}");
        tmplBuilder.EndRow();
        tmplBuilder.EndTable();

        // Save the template (required by the task description).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and clone its table.
        // -----------------------------------------------------------------
        Document srcDoc = new Document(templatePath);
        Table sourceTable = srcDoc.FirstSection.Body.Tables[0];
        Table clonedTable = (Table)sourceTable.Clone(true); // deep clone, still belongs to srcDoc

        // -----------------------------------------------------------------
        // 4. Create the result document and import the cloned table.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();

        // Import the cloned table so that it belongs to resultDoc.
        Table importedTable = (Table)resultDoc.ImportNode(clonedTable, true);
        resultDoc.FirstSection.Body.AppendChild(importedTable);

        // -----------------------------------------------------------------
        // 5. Replace placeholders inside the imported table.
        // -----------------------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        importedTable.Range.Replace("{{Name}}",    "John Doe",   replaceOptions);
        importedTable.Range.Replace("{{Age}}",     "30",         replaceOptions);
        importedTable.Range.Replace("{{City}}",    "New York",   replaceOptions);
        importedTable.Range.Replace("{{Country}}", "USA",        replaceOptions);

        // -----------------------------------------------------------------
        // 6. Save the resulting document.
        // -----------------------------------------------------------------
        resultDoc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
