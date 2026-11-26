using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style.
        const string customStyleName = "MyCustomStyle";
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, customStyleName);
        // Set style properties (avoid System.Drawing types).
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Add sample paragraphs containing the target text.
        builder.Writeln("This is the first paragraph with PLACEHOLDER text.");
        builder.Writeln("Second paragraph without the keyword.");
        builder.Writeln("Third paragraph also has PLACEHOLDER inside.");
        builder.Writeln("PLACEHOLDER at the start of a paragraph.");
        builder.Writeln("End of document.");

        // Set up find‑replace with a callback that applies the custom style.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new StyleApplyingCallback(customStyleName);

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces the matched text and applies a custom style to the containing paragraph.
    private class StyleApplyingCallback : IReplacingCallback
    {
        private readonly string _styleName;

        public StyleApplyingCallback(string styleName)
        {
            _styleName = styleName;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace the matched text.
            args.Replacement = "REPLACED";

            // Locate the paragraph that contains the match.
            Paragraph? paragraph = args.MatchNode?.ParentNode?.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph != null)
            {
                // Use DocumentBuilder to apply the custom style to the paragraph.
                DocumentBuilder builder = new DocumentBuilder((Document)paragraph.Document);
                builder.MoveTo(paragraph);
                builder.ParagraphFormat.Style = builder.Document.Styles[_styleName];
            }

            return ReplaceAction.Replace;
        }
    }
}
