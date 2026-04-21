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

        // Configure paragraph indents.
        double paragraphLeftIndent = 50;   // points
        double paragraphRightIndent = 30;  // points
        builder.ParagraphFormat.LeftIndent = paragraphLeftIndent;
        builder.ParagraphFormat.RightIndent = paragraphRightIndent;

        // Write a sample paragraph.
        builder.Writeln("This paragraph has custom left and right indents.");

        // Start a table.
        Table table = builder.StartTable();

        // Add a single cell with some text.
        builder.InsertCell();
        builder.Write("Table content");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = paragraphLeftIndent;

        // NOTE: Aspose.Words does not provide a Table.RightIndent property.
        // Therefore, only the left indent can be aligned directly.
        // Right alignment can be achieved indirectly by adjusting the table width
        // or using other layout options if needed.

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableIndent.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        // The program ends automatically; no user interaction required.
    }
}
