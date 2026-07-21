using System;
using System.Drawing;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph with a PLACEHOLDER.");
        builder.Writeln("Second paragraph also contains PLACEHOLDER text.");
        builder.Writeln("No placeholder here.");
        builder.Writeln("Another PLACEHOLDER appears in this line.");

        // Define a custom style that will be applied to replaced paragraphs.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Color = Color.Blue;
        customStyle.Font.Size = 14;

        // Set up find‑replace options with a callback that applies the style.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ReplaceAndStyleCallback(customStyle.Name)
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "Replaced", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces the matched text and applies the custom style to the containing paragraph.
    private class ReplaceAndStyleCallback : IReplacingCallback
    {
        private readonly string _styleName;

        public ReplaceAndStyleCallback(string styleName)
        {
            _styleName = styleName ?? throw new ArgumentNullException(nameof(styleName));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Set the replacement text.
            args.Replacement = "Replaced";

            // Find the paragraph that contains the match and apply the style.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
                paragraph.ParagraphFormat.StyleName = _styleName;

            return ReplaceAction.Replace;
        }
    }
}
