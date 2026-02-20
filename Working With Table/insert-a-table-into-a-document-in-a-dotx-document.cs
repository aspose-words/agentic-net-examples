using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDotx
{
    static void Main()
    {
        // Path to the folder that contains the DOTX template.
        string dataDir = @"C:\Docs\";

        // Load the DOTX template.
        string templatePath = Path.Combine(dataDir, "Template.dotx");
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table at the current cursor position.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Write("Value 1");
        builder.InsertCell();
        builder.Write("Value 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style and auto‑fit the table to its contents.
        table.StyleIdentifier = StyleIdentifier.LightListAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document.
        string outputPath = Path.Combine(dataDir, "Result.docx");
        doc.Save(outputPath);
    }
}
