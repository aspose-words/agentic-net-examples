using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Path to the source DOTM template.
        string dataDir = @"C:\Data\";
        string templatePath = System.IO.Path.Combine(dataDir, "Template.dotm");
        string outputPath = System.IO.Path.Combine(dataDir, "Result.docx");

        // Load the DOTM document.
        Document doc = new Document(templatePath);

        // Find the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Ensure the table has at least one row before applying style options.
        table.EnsureMinimum();

        // Optionally set a built‑in style identifier.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply desired style options (e.g., first column, first row, and row banding).
        table.StyleOptions = TableStyleOptions.FirstColumn |
                             TableStyleOptions.FirstRow |
                             TableStyleOptions.RowBands;

        // Optionally auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document.
        doc.Save(outputPath);
    }
}
