using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample paragraphs containing the text to be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with PLACEHOLDER text.");
        builder.Writeln("Second paragraph also has PLACEHOLDER inside.");
        builder.Writeln("Third paragraph without the keyword.");

        // Define a custom paragraph style that will be applied to each replaced paragraph.
        const string customStyleName = "MyCustomStyle";
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, customStyleName);
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Set up find-and-replace with a callback that applies the custom style.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ParagraphStyleCallback(doc, customStyleName)
        };

        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces the matched text and applies the custom style to the containing paragraph.
    private class ParagraphStyleCallback : IReplacingCallback
    {
        private readonly Document _document;
        private readonly string _styleName;

        public ParagraphStyleCallback(Document document, string styleName)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _styleName = styleName ?? throw new ArgumentNullException(nameof(styleName));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace the found text.
            args.Replacement = "REPLACED";

            // Locate the paragraph that contains the match.
            Paragraph paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph != null)
            {
                // Use DocumentBuilder to move to the paragraph and apply the custom style.
                DocumentBuilder builder = new DocumentBuilder(_document);
                builder.MoveTo(paragraph);
                builder.ParagraphFormat.StyleName = _styleName;
            }

            return ReplaceAction.Replace;
        }
    }
}
