using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field at the beginning of the document.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a blank line after the TOC.
        builder.Writeln();

        // Add Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // Add normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is some introductory text.");

        // Add Heading 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // Add a numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.ListFormat.RemoveNumbers();

        // Add another Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");

        // Update all fields (including the TOC) to reflect the new content.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "TableOfContents.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
