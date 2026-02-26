using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a built‑in style (Heading1) to the first part of the line.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator – creates a new paragraph without a line break.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
        customStyle.Font.Bold = false;
        customStyle.Font.Size = 8;
        customStyle.Font.Name = "Arial";

        // Apply the custom style to the second part of the line.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Verify that the separator created a separate paragraph.
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        Paragraph secondPara = doc.FirstSection.Body.Paragraphs[1];
        Console.WriteLine("First paragraph style: " + firstPara.ParagraphFormat.Style.Name);
        Console.WriteLine("Second paragraph style: " + secondPara.ParagraphFormat.Style.Name);
        Console.WriteLine("First paragraph is a style separator: " + firstPara.BreakIsStyleSeparator);
        Console.WriteLine("Second paragraph is a style separator: " + secondPara.BreakIsStyleSeparator);

        // Save the document.
        doc.Save("StyleSeparator.docx");
    }
}
