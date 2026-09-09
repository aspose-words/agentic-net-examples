using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the source DOCX and the resulting HTML
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultPath = Path.Combine(outputDir, "Result.html");

        // Create a source document with sample content
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        srcBuilder.Writeln("It will be inserted into another document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Create the destination document
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document start.");

        // Insert a page break before the inserted content
        destBuilder.InsertBreak(BreakType.PageBreak);

        // Load the source document to be inserted
        Document srcToInsert = new Document(sourcePath);

        // Insert the source document at the current cursor position, preserving its formatting
        destBuilder.InsertDocument(srcToInsert, ImportFormatMode.KeepSourceFormatting);
        destBuilder.Writeln("Destination document end.");

        // Save the combined document as HTML
        destDoc.Save(resultPath, SaveFormat.Html);

        // Verify that the HTML file was created
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {resultPath}");
        }
    }
}
