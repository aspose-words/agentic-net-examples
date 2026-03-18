using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    // Directory for generated files (artifacts).
    private static readonly string ArtifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");

    static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(ArtifactsDir);

        // Run the test logic.
        ExtractParagraphs_ShouldRetainOriginalStyling();

        Console.WriteLine("Test passed.");
    }

    private static void ExtractParagraphs_ShouldRetainOriginalStyling()
    {
        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Create a custom paragraph style with distinct formatting.
        Style customStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        customStyle.Font.Color = Color.Blue;

        // Paragraph 1 – uses the custom style.
        srcBuilder.ParagraphFormat.Style = customStyle;
        srcBuilder.Writeln("This is styled text.");

        // Paragraph 2 – normal style (acts as the end delimiter).
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("This is normal text.");

        // Save the source document (optional, for inspection).
        srcDoc.Save(Path.Combine(ArtifactsDir, "Source.docx"));

        // ---------- Identify range to extract ----------
        // For this example we extract from the first paragraph up to (and including) the second paragraph.
        ParagraphCollection paragraphs = srcDoc.FirstSection.Body.Paragraphs;
        int startIndex = 0; // first paragraph
        int endIndex = 1;   // second paragraph

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Import each paragraph from the source into the destination,
        // preserving the original formatting.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = paragraphs[i];
            Node importedNode = destDoc.ImportNode(srcParagraph, true, ImportFormatMode.KeepSourceFormatting);
            destBuilder.InsertNode(importedNode);
        }

        // Save the extracted document.
        string extractedPath = Path.Combine(ArtifactsDir, "Extracted.docx");
        destDoc.Save(extractedPath);

        // ---------- Verify styling is retained ----------
        Paragraph extractedPara1 = destDoc.FirstSection.Body.Paragraphs[0];
        Paragraph extractedPara2 = destDoc.FirstSection.Body.Paragraphs[1];

        // Verify the first paragraph kept the custom style name.
        if (extractedPara1.ParagraphFormat.StyleName != "MyCustomStyle")
            throw new Exception($"Expected style name 'MyCustomStyle' but got '{extractedPara1.ParagraphFormat.StyleName}'.");

        // Verify the run inside the first paragraph retained font properties.
        Run run = (Run)extractedPara1.Runs[0];
        if (run.Font.Name != "Arial")
            throw new Exception($"Expected font name 'Arial' but got '{run.Font.Name}'.");
        if (run.Font.Size != 14)
            throw new Exception($"Expected font size 14 but got {run.Font.Size}.");
        if (!run.Font.Bold)
            throw new Exception("Expected font to be bold.");
        if (run.Font.Color.ToArgb() != Color.Blue.ToArgb())
            throw new Exception("Expected font color to be Blue.");

        // Verify the second paragraph uses the Normal style.
        if (extractedPara2.ParagraphFormat.StyleIdentifier != StyleIdentifier.Normal)
            throw new Exception("Second paragraph does not use Normal style.");
    }
}
