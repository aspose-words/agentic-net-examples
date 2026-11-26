using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with heading styles to represent chapters.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the introduction chapter content.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This chapter provides detailed information.");

        // Save the Word document as PDF (input file for conversion).
        string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The input PDF was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to EPUB while preserving headings as navigation.
        string epubPath = "output.epub";
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Verify the EPUB was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The output EPUB was not created.");

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine("PDF successfully converted to EPUB.");
    }
}
