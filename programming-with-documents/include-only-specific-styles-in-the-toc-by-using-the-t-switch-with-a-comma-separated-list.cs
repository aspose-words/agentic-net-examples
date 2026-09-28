using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title using the built‑in Title style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Document Title");

        // Add headings with different built‑in heading styles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1 – Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1 – Heading 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1 – Heading 3");

        // Add a normal paragraph (not a heading).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Regular paragraph, not a heading.");

        // Insert a Table of Contents that includes only Heading 1 and Heading 3.
        // Use the \\t switch with a comma‑separated list of style names.
        // The InsertField overload that takes a field code string is used to avoid overload ambiguity.
        builder.InsertField("TOC \\o \"1-3\" \\h \\z \\t \"Heading 1,Heading 3\"");

        // Update fields so the TOC is generated.
        doc.UpdateFields();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TOC_Styles.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
