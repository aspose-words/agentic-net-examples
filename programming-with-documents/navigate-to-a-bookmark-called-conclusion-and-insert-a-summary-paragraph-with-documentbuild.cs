using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string sourceFile = "Sample.docx";
        const string outputFile = "Output.docx";

        // -------------------------------------------------
        // Step 1: Create a sample document with a bookmark
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content
        builder.Writeln("This is the beginning of the document.");
        builder.Writeln();

        // Insert a bookmark named "Conclusion"
        builder.StartBookmark("Conclusion");
        builder.Writeln("Conclusion placeholder text.");
        builder.EndBookmark("Conclusion");

        // Save the source document
        doc.Save(sourceFile);

        // -------------------------------------------------
        // Step 2: Load the document and navigate to the bookmark
        // -------------------------------------------------
        Document loadedDoc = new Document(sourceFile);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // Move to the bookmark "Conclusion"
        loadedBuilder.MoveToBookmark("Conclusion");

        // Insert a new paragraph after the bookmark and add the summary text
        loadedBuilder.InsertParagraph();
        loadedBuilder.Writeln("This is the summary paragraph inserted after the Conclusion bookmark.");

        // Save the modified document
        loadedDoc.Save(outputFile);

        // -------------------------------------------------
        // Optional: Verify that the output file was created
        // -------------------------------------------------
        if (File.Exists(outputFile))
        {
            Console.WriteLine($"Document saved successfully to '{outputFile}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the output document.");
        }
    }
}
