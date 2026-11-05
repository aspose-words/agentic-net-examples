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

        // Add some paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Target paragraph."); // This is the paragraph after which we will insert the table.
        builder.Writeln("Third paragraph.");

        // Locate the paragraph with the exact text "Target paragraph."
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Trim() == "Target paragraph.")
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new Exception("Target paragraph not found.");

        // Move the builder to the target paragraph.
        builder.MoveTo(targetParagraph);
        // Insert a new empty paragraph after the target paragraph – this will be the anchor for the table.
        builder.InsertParagraph();

        // Build a simple 2x2 table using DocumentBuilder (preferred workflow).
        builder.StartTable();
        for (int rowIndex = 0; rowIndex < 2; rowIndex++)
        {
            for (int colIndex = 0; colIndex < 2; colIndex++)
            {
                builder.InsertCell();
                builder.Write($"R{rowIndex + 1}C{colIndex + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the table that was just created (it will be the last table in the document).
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true).Last();

        // Apply a basic grid style now that the table has rows.
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the output document.");
    }
}
