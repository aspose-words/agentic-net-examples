using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a target paragraph one.");
        builder.Writeln("Another line without the keyword.");
        builder.Writeln("Second target paragraph is here.");
        builder.Writeln("target appears multiple times in this target paragraph.");

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Color = System.Drawing.Color.Blue; // Use fully qualified System.Drawing type without importing the namespace.
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Set up find‑replace with a callback that applies the custom style.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new StyleApplyingCallback(doc, customStyle);

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("target", "replaced", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were made.");

        // Save the modified document.
        doc.Save("styled-output.docx");
    }
}

// Callback that applies the custom style to each paragraph containing a replacement.
class StyleApplyingCallback : IReplacingCallback
{
    private readonly Document _document;
    private readonly Style _style;

    public StyleApplyingCallback(Document document, Style style)
    {
        _document = document;
        _style = style;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        Node matchNode = args.MatchNode;
        if (matchNode != null)
        {
            Paragraph paragraph = matchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph != null)
            {
                // Apply the custom style to the whole paragraph.
                paragraph.ParagraphFormat.Style = _style;
            }
        }
        return ReplaceAction.Replace;
    }
}
