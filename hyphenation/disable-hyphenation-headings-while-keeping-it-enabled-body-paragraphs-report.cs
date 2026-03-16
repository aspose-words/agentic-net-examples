using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a body paragraph – hyphenation will be enabled for this text.
        builder.Font.Size = 12;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Add a heading – we will later suppress hyphenation for this paragraph.
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.Writeln("Chapter 1 – Introduction");

        // Add another body paragraph.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Suppress hyphenation for all heading paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // IsHeading is true for built‑in heading styles (Heading 1, Heading 2, …).
            if (para.ParagraphFormat.IsHeading)
            {
                para.ParagraphFormat.SuppressAutoHyphens = true;
            }
        }

        // Save the document.
        doc.Save("ReportWithSelectiveHyphenation.docx");
    }
}
