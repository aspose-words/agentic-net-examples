using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the paragraph with Heading1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("First part ");

        // Insert a style separator run.
        // This separates the paragraph into two runs that can have different styles.
        builder.InsertStyleSeparator();

        // Second part of the paragraph with Normal style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Write("second part.");

        // Save the document to disk.
        doc.Save("StyleSeparator.docx");
    }
}
