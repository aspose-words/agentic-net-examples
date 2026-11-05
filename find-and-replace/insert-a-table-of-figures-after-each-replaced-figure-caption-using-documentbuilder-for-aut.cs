using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for inserting sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs with figure captions.
        builder.Writeln("Sample Document");
        builder.Writeln("Figure 1: Old Caption");
        builder.Writeln("Some intermediate text.");
        builder.Writeln("Figure 2: Old Caption");

        // Set up find‑and‑replace options with a custom callback that inserts a Table of Figures.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new TableOfFiguresInserter(doc);

        // Replace the old caption text.
        int replacedCount = doc.Range.Replace("Old Caption", "New Caption", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Update fields so that the inserted Table of Figures reflects the current captions.
        doc.UpdateFields();

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Callback that inserts a Table of Figures after each replaced figure caption.
public class TableOfFiguresInserter : IReplacingCallback
{
    private readonly Document _document;

    public TableOfFiguresInserter(Document document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Locate the paragraph that contains the match.
        Node matchNode = args.MatchNode;
        Paragraph paragraph = matchNode?.ParentNode?.GetAncestor(NodeType.Paragraph) as Paragraph;

        if (paragraph != null)
        {
            // Use a builder positioned after the found paragraph.
            DocumentBuilder builder = new DocumentBuilder(_document);
            builder.MoveTo(paragraph);
            builder.InsertParagraph(); // Move cursor to a new paragraph after the original one.

            // Insert a Table of Figures for the "Figure" label using a TOC field.
            // The field code creates a table of figures based on the "Figure" label.
            builder.InsertField(@"TOC \h \z \c ""Figure""");
            builder.Writeln(); // Ensure there is a line break after the field.
        }

        // Continue with the normal replacement.
        return ReplaceAction.Replace;
    }
}
