using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with figure captions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Introduction paragraph.");
        builder.Writeln("Figure 1: Old caption for first figure.");
        builder.Writeln("Some text between figures.");
        builder.Writeln("Figure 2: Old caption for second figure.");
        builder.Writeln("Conclusion paragraph.");

        // Save the initial document (optional, just for reference).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define a regex to find figure captions like "Figure 1: ..."
        Regex figureCaptionRegex = new Regex(@"Figure (\d+): .+", RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new CaptionReplacer();

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace(figureCaptionRegex, "", options);
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one figure caption replacement.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }

    private class CaptionReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // args.Match is a System.Text.RegularExpressions.Match.
            Match match = args.Match;
            string figureNumber = match.Groups[1].Value;

            // Define the new caption text.
            string newCaption = $"Figure {figureNumber}: Updated caption.";

            // Set the replacement text.
            args.Replacement = newCaption;

            // Insert a Table of Figures after the paragraph containing the match.
            Node matchNode = args.MatchNode;
            Paragraph paragraph = (Paragraph)matchNode.GetAncestor(NodeType.Paragraph);
            if (paragraph != null)
            {
                // The document associated with the node is a Document (not just DocumentBase).
                Document doc = (Document)matchNode.Document;
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Move the builder to the paragraph that contains the match.
                builder.MoveTo(paragraph);
                // Write a new empty paragraph to separate the caption from the table.
                builder.Writeln();

                // Insert a Table of Figures field.
                // The field code "TOC \\h \\z \\c \"Figure\"" creates a table of figures.
                builder.InsertField("TOC \\h \\z \\c \"Figure\"");
                builder.Writeln();
            }

            return ReplaceAction.Replace;
        }
    }
}
