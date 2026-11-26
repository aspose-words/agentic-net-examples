using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style for TOC entries.
        Style tocStyle = doc.Styles.Add(StyleType.Paragraph, "MyTOCStyle");
        tocStyle.Font.Name = "Arial";
        tocStyle.Font.Size = 12;
        tocStyle.Font.Color = Color.Blue;
        tocStyle.Font.Bold = true;

        // Insert a Table of Contents field and specify the custom style using the \\t switch.
        // The switch format: \t "StyleName"
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u \\t \"MyTOCStyle\"");

        // Add headings that will be captured by the TOC.
        builder.Font.Size = 16;
        builder.Font.Bold = true;

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.Writeln("Content of chapter 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Content of section 1.1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.Writeln("Content of chapter 2.");

        // Update fields so the TOC reflects the new style.
        doc.UpdateFields();

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
