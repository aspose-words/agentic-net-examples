using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing; // For Aspose.Words drawing types if needed

public class Program
{
    public static void Main()
    {
        // Create a blank document and add some paragraphs containing the word "old".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first old paragraph.");
        builder.Writeln("Another old paragraph follows.");
        builder.Writeln("No match here.");
        builder.Writeln("The last old paragraph.");

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        // Font color is optional; omitted to avoid System.Drawing usage.
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Set up find-and-replace with a callback that applies the custom style.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ParagraphStyler();

        int replacedCount = doc.Range.Replace("old", "new", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    private class ParagraphStyler : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Apply the custom style to the paragraph that contains the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // Cast DocumentBase to Document for the builder.
                DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);
                cb.MoveTo(paragraph);
                cb.ParagraphFormat.StyleName = "MyCustomStyle";
            }

            // Perform the standard replacement.
            args.Replacement = args.Match.Value.Replace("old", "new");
            return ReplaceAction.Replace;
        }
    }
}
