using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents placeholder at the beginning of the document.
        builder.MoveToDocumentStart();
        // The TOC will include outline levels 1 through 3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a main heading (outline level 1) using the built‑in Heading1 style.
        builder.MoveToDocumentEnd();
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Main Heading");

        // Add a subheading using the built‑in Heading2 style (outline level 2).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Subheading (Heading2)");

        // Add a custom paragraph that will act as a subheading.
        // First set it to the Normal style, then manually set its outline level to 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Custom Subheading");
        Paragraph customPara = builder.CurrentParagraph;
        customPara.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;

        // Save the document to the output file.
        doc.Save("OutputDocument.docx");
    }
}
