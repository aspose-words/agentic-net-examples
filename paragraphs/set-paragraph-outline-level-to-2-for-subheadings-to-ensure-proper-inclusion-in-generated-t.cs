using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Document Title");

        // Insert a level‑1 heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Insert a level‑2 subheading and explicitly set its outline level to 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Section 1.1");

        // Another level‑2 subheading.
        builder.Writeln("Section 1.2");

        // Insert another level‑1 heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1; // reset to default for heading 1
        builder.Writeln("Chapter 2");

        // Level‑2 subheading under Chapter 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Section 2.1");

        // Insert a Table of Contents at the beginning of the document.
        builder.MoveToDocumentStart();
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Save the document.
        doc.Save("Output.docx");
    }
}
