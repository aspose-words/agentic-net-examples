using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;
using System.Drawing;

class ApplyStyleCallback : IReplacingCallback
{
    private readonly string _styleName;

    public ApplyStyleCallback(string styleName)
    {
        _styleName = styleName;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Apply the custom paragraph style to the paragraph that contains the match.
        Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;
        paragraph.ParagraphFormat.StyleName = _styleName;

        // Keep the original replacement text unchanged.
        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Create a new document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample content with placeholders to be replaced.
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Thank you for your purchase.");
        builder.Writeln("Best regards,");
        builder.Writeln("_Company_");

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 12;
        customStyle.Font.Color = Color.DarkBlue;
        customStyle.ParagraphFormat.SpaceAfter = 10;
        customStyle.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Set up find/replace options with a callback that applies the style.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ApplyStyleCallback(customStyle.Name);

        // Perform replacements; each paragraph that contains a match will receive the custom style.
        doc.Range.Replace("_Customer_", "John Doe", options);
        doc.Range.Replace("_Company_", "Acme Corp.", options);

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
