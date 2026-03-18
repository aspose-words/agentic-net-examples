using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SetOutlineLevelExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field that will pick up paragraphs with outline levels.
        // The switches specify that headings of levels 1‑3 are included.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // ----- Main heading (level 1) -----
        // Use the built‑in Heading 1 style; this automatically sets OutlineLevel to Level1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // ----- Subheading (level 2) -----
        // Use a normal style but explicitly set the outline level to Level2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("1.1 Overview");

        // Another subheading.
        builder.Writeln("1.2 Details");

        // ----- Another main heading (level 1) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Usage");

        // Subheading under the second chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("2.1 Getting Started");

        // Update all fields (including the TOC) to reflect the new outline levels.
        doc.UpdateFields();

        // Save the document.
        doc.Save("OutlineLevelExample.docx", SaveFormat.Docx);
    }
}
