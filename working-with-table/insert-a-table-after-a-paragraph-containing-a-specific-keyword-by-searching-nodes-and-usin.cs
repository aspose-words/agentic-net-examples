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

        // Add some paragraphs, one of them contains the keyword.
        builder.Writeln("First paragraph.");
        builder.Writeln("Paragraph with keyword INSERT_HERE where table will be inserted.");
        builder.Writeln("Last paragraph.");

        // Search for the paragraph that contains the specific keyword.
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Contains("INSERT_HERE"))
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Keyword not found in any paragraph.");

        // Insert an empty paragraph after the target paragraph – this will be the anchor for the new table.
        CompositeNode parent = targetParagraph.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Unable to locate a suitable parent node for insertion.");

        Paragraph anchorParagraph = new Paragraph(doc);
        parent.InsertAfter(anchorParagraph, targetParagraph);

        // Move the builder to the newly inserted paragraph and build a 2x2 table.
        builder.MoveTo(anchorParagraph);
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
