using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several headings that will be replaced later.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.Writeln("Heading 2");
        builder.Writeln("Heading 3");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertPageNumberCallback();

        // Replace the word "Heading" with "Section". The callback will insert a PAGE field after each replaced heading.
        int replaced = doc.Range.Replace("Heading", "Section", options);
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Update fields so that PAGE fields display correct page numbers.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a PAGE field after the paragraph containing the match.
    private class InsertPageNumberCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace the matched text with the new value.
            args.Replacement = "Section";

            // Locate the paragraph that contains the match.
            Paragraph headingParagraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
            if (headingParagraph == null)
                return ReplaceAction.Replace;

            // Use a DocumentBuilder positioned at the document that owns the paragraph.
            DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);
            cb.MoveTo(headingParagraph);
            // Insert a new paragraph after the heading and place a PAGE field there.
            cb.InsertParagraph(); // Cursor now in the new paragraph.
            cb.InsertField(FieldType.FieldPage, true);

            return ReplaceAction.Replace;
        }
    }
}
