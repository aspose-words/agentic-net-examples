using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("This is the introductory paragraph before any headings.");

        // Define headings to add.
        string[] headings = { "Chapter 1: Getting Started", "Chapter 2: Advanced Topics", "Chapter 3: Conclusion" };

        foreach (string heading in headings)
        {
            // Insert a page break before the heading.
            builder.InsertBreak(BreakType.PageBreak);

            // Apply the Heading 1 style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // Write the heading text.
            builder.Writeln(heading);

            // Reset style to Normal for subsequent content (optional).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
