using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a folder for the output document.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Initialize a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define two custom paragraph styles that will be referenced by the TOC.
        Style customStyle1 = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle1");
        customStyle1.Font.Size = 16;
        customStyle1.Font.Bold = true;

        Style customStyle2 = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle2");
        customStyle2.Font.Size = 14;
        customStyle2.Font.Italic = true;

        // Insert a Table of Contents that includes only the custom styles.
        // The \t switch takes a comma‑separated list of "StyleName; TOCLevel".
        // The \h switch makes the entries clickable hyperlinks.
        string tocSwitches = @"\t ""MyCustomStyle1;1, MyCustomStyle2;2"" \h";
        builder.InsertTableOfContents(tocSwitches);
        builder.InsertBreak(BreakType.PageBreak);

        // Add content using the custom styles – these entries will appear in the TOC.
        builder.ParagraphFormat.Style = customStyle1;
        builder.Writeln("Custom Heading 1");

        builder.ParagraphFormat.Style = customStyle2;
        builder.Writeln("Custom Subheading 1");

        // Add a paragraph with a built‑in heading style to demonstrate it is NOT included.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Built‑in Heading (should not appear in TOC)");

        // Update fields so the TOC reflects the added entries.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save(Path.Combine(artifactsDir, "TOC_CustomStyles.docx"));
    }
}
