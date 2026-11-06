using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample headings and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is some introductory text.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details about section 1.1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("More content follows.");

        // Collect all heading paragraphs.
        List<Paragraph> headings = new List<Paragraph>();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier style = para.ParagraphFormat.StyleIdentifier;
            if (style >= StyleIdentifier.Heading1 && style <= StyleIdentifier.Heading9)
                headings.Add(para);
        }

        // Insert a table after each heading.
        foreach (Paragraph heading in headings)
        {
            Table table = CreateSampleTable(doc);
            // Insert the table after the heading paragraph.
            heading.ParentNode.InsertAfter(table, heading);
        }

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");
    }

    // Creates a simple 2x2 table using DocumentBuilder on a temporary document,
    // then imports it into the target document.
    private static Table CreateSampleTable(Document targetDoc)
    {
        // Build the table in a temporary document.
        Document tempDoc = new Document();
        DocumentBuilder tb = new DocumentBuilder(tempDoc);
        tb.StartTable();

        tb.InsertCell();
        tb.Write("Cell 1");
        tb.InsertCell();
        tb.Write("Cell 2");
        tb.EndRow();

        tb.InsertCell();
        tb.Write("Cell 3");
        tb.InsertCell();
        tb.Write("Cell 4");
        tb.EndRow();

        tb.EndTable();

        // Retrieve the created table.
        Table tempTable = (Table)tb.CurrentParagraph.PreviousSibling;

        // Import the table into the target document.
        NodeImporter importer = new NodeImporter(tempDoc, targetDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(tempTable, true);

        return importedTable;
    }
}
