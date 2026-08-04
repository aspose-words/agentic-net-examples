using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define paragraph indentation that we want the table to align with.
        double paragraphLeftIndent = 50.0;   // points
        double paragraphRightIndent = 30.0;  // points

        // Apply the indentation to the builder's paragraph format.
        builder.ParagraphFormat.LeftIndent = paragraphLeftIndent;
        builder.ParagraphFormat.RightIndent = paragraphRightIndent;

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = paragraphLeftIndent;

        // NOTE: Table.RightIndent is not available per the library restrictions,
        // so we rely on the paragraph's right indent for visual alignment.

        // Add some sample content to the cell.
        builder.Writeln("This table aligns its left edge with the paragraph's left indent.");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a local file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "TableIndentAlignment.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
