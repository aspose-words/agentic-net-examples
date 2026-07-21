using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace AsposeWordsFindReplaceExample
{
    // Callback that replaces the matched text and inserts a Table of Figures field after the paragraph.
    public class FigureCaptionReplacer : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Replace the matched text with the new caption fragment.
            args.Replacement = "New";

            // The paragraph that contains the matched text.
            Paragraph captionParagraph = args.MatchNode.ParentNode as Paragraph;
            if (captionParagraph == null)
                return ReplaceAction.Skip;

            // The document that owns the paragraph.
            Document doc = (Document)captionParagraph.Document;

            // Create a new empty paragraph that will hold the Table of Figures field.
            Paragraph tocParagraph = new Paragraph(doc);
            CompositeNode parent = captionParagraph.ParentNode;
            parent.InsertAfter(tocParagraph, captionParagraph);

            // Insert the TOC field with switches for a Table of Figures (entries labeled "Figure").
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(tocParagraph);
            builder.InsertField("TOC \\h \\z \\c \"Figure\"");

            // Continue with the normal replacement of the matched text.
            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample figure captions.
            builder.Writeln("Figure 1: Old Caption");
            builder.Writeln("Some introductory text.");
            builder.Writeln("Figure 2: Another Old Caption");
            builder.Writeln("Additional content.");

            // Set up find-and-replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions(new FigureCaptionReplacer());

            // Replace the word "Old" in figure captions with "New" and insert a Table of Figures after each.
            int replacedCount = doc.Range.Replace("Old", "New", options);
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one replacement.");

            // Update all fields so that the inserted Table of Figures fields display correctly.
            doc.UpdateFields();

            // Save the resulting document.
            doc.Save("output.docx");
        }
    }
}
