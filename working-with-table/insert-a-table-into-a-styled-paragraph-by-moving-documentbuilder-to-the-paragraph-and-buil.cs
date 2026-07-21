using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph and apply a built‑in style (Heading1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Styled Paragraph");

        // Keep a reference to the paragraph we just created.
        Paragraph styledParagraph = builder.CurrentParagraph;

        // Move the builder's cursor to the styled paragraph.
        builder.MoveTo(styledParagraph);

        // Build a 2x2 table at the current position.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableInStyledParagraph.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
