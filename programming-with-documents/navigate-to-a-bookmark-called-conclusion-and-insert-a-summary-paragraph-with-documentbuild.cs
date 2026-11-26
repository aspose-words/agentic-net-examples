using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing a bookmark named "Conclusion".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the introduction.");
        builder.StartBookmark("Conclusion");
        builder.Writeln("This is the conclusion placeholder.");
        builder.EndBookmark("Conclusion");
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Load the document and navigate to the "Conclusion" bookmark.
        Document loadedDoc = new Document(sourcePath);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);
        loadedBuilder.MoveToBookmark("Conclusion");

        // Insert a summary paragraph at the bookmark location.
        loadedBuilder.Writeln("Summary: This document demonstrates inserting a paragraph at a bookmark.");

        // Save the modified document.
        string outputPath = "Output.docx";
        loadedDoc.Save(outputPath);
    }
}
