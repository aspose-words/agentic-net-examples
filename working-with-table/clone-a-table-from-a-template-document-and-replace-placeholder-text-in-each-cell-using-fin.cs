using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the generated files
        string outputPath = "ClonedTable.docx";

        // -------------------------------------------------
        // 1. Create a template document that contains a table
        //    with placeholder text in each cell.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

        // Build a 2x2 table with placeholders.
        tmplBuilder.StartTable();

        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{FirstName}}");   // Cell (0,0)

        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{LastName}}");    // Cell (0,1)

        tmplBuilder.EndRow();

        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Age}}");         // Cell (1,0)

        tmplBuilder.InsertCell();
        tmplBuilder.Write("{{Country}}");    // Cell (1,1)

        tmplBuilder.EndTable();

        // -------------------------------------------------
        // 2. Retrieve the table from the template document.
        // -------------------------------------------------
        Table templateTable = (Table)templateDoc.GetChildNodes(NodeType.Table, true)[0];

        // -------------------------------------------------
        // 3. Clone the table node (deep clone).
        // -------------------------------------------------
        Table clonedTable = (Table)templateTable.Clone(true);

        // -------------------------------------------------
        // 4. Create the destination document where the cloned
        //    table will be inserted.
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Cloned Table Below:");

        // -------------------------------------------------
        // 5. Import the cloned table into the destination
        //    document (required because the table belongs to
        //    the template document).
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(templateDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(clonedTable, true);
        destDoc.FirstSection.Body.AppendChild(importedTable);

        // -------------------------------------------------
        // 6. Replace placeholder text in the cloned table using
        //    FindReplaceOptions.
        // -------------------------------------------------
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        destDoc.Range.Replace("{{FirstName}}", "John", replaceOptions);
        destDoc.Range.Replace("{{LastName}}", "Doe", replaceOptions);
        destDoc.Range.Replace("{{Age}}", "30", replaceOptions);
        destDoc.Range.Replace("{{Country}}", "USA", replaceOptions);

        // -------------------------------------------------
        // 7. Save the resulting document.
        // -------------------------------------------------
        destDoc.Save(outputPath);

        // -------------------------------------------------
        // 8. Simple validation to ensure the file was created.
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // The program finishes here without any interactive prompts.
    }
}
