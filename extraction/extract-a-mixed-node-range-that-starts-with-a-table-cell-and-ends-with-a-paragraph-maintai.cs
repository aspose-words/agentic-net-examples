using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractMixedRange
{
    public static void Main()
    {
        // Create a source document with mixed content: a table and a paragraph.
        Document srcDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(srcDoc);

        // Paragraph before the range (will not be extracted).
        builder.Writeln("Paragraph before the extracted range.");

        // Build a table. The first cell will be the start of the range.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Start cell (range begins here).");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Second cell.");
        builder.EndRow();

        builder.EndTable();

        // Paragraph that will be the end of the range.
        Paragraph endParagraph = new Paragraph(srcDoc);
        endParagraph.AppendChild(new Run(srcDoc, "End paragraph of the extracted range."));
        srcDoc.FirstSection.Body.AppendChild(endParagraph);

        // Additional content after the range (will not be extracted).
        builder.Writeln("Paragraph after the extracted range.");

        // -----------------------------------------------------------------
        // Extraction: copy the table (starting from its first cell) and the
        // ending paragraph into a new document while preserving layout.
        // -----------------------------------------------------------------

        // Destination document – start with an empty structure.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // remove the default section/paragraph
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Use NodeImporter to import nodes with source formatting.
        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the whole table (its first cell is the start of the range).
        Node importedTable = importer.ImportNode(table, true);
        destBody.AppendChild(importedTable);

        // Import the ending paragraph.
        Node importedEndParagraph = importer.ImportNode(endParagraph, true);
        destBody.AppendChild(importedEndParagraph);

        // Save the extracted content.
        const string outputPath = "ExtractedRange.docx";
        destDoc.Save(outputPath);
        Console.WriteLine($"Extracted range saved to '{outputPath}'.");
    }
}
