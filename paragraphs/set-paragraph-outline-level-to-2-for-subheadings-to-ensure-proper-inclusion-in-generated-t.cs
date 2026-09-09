using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that will include headings up to level 3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a main heading (outline level 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Main Heading");

        // Insert a subheading and explicitly set its outline level to 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2; // Level 2 = outline level 2
        builder.Writeln("Subheading Level 2");

        // Add a normal paragraph of body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is some body text under the subheading.");

        // Update fields so the TOC reflects the headings.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("ParagraphOutlineLevelExample.docx");
    }
}
