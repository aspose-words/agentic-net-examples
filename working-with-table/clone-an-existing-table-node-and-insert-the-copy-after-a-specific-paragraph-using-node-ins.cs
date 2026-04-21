using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class CloneTableExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will serve as the insertion point.
        builder.Writeln("Paragraph before the original table.");

        // Build a simple 2‑cell table.
        Table originalTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add another paragraph after the original table.
        builder.Writeln("Paragraph after the original table.");

        // Locate the paragraph after which we want to insert the cloned table.
        // Here we use the first paragraph in the document (index 0).
        Paragraph referenceParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Clone the original table (deep clone).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Insert the cloned table after the reference paragraph.
        // The parent node of a paragraph is the Body of the section.
        Body body = referenceParagraph.ParentNode as Body;
        if (body == null)
            throw new InvalidOperationException("Unable to locate the body of the document.");

        body.InsertAfter(clonedTable, referenceParagraph);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ClonedTable.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
    }
}
