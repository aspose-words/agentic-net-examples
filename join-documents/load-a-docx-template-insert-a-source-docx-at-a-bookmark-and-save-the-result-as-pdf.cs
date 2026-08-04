using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document with a bookmark where the source will be inserted.
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Template start");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("Placeholder text (will be replaced)");
        templateBuilder.EndBookmark("InsertHere");
        templateBuilder.Writeln("Template end");

        // Create the source document that will be inserted.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the inserted content from the source document.");

        // Move to the bookmark in the template and insert the source document.
        templateBuilder.MoveToBookmark("InsertHere");
        templateBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Define the output PDF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.pdf");

        // Save the merged document as PDF.
        templateDoc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created and contains expected text.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output PDF was not created.", outputPath);

        // Load the PDF back as a document to verify its text content.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("Template start") ||
            !resultText.Contains("This is the inserted content from the source document.") ||
            !resultText.Contains("Template end"))
        {
            throw new InvalidOperationException("The merged PDF does not contain expected content.");
        }
    }
}
