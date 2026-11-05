using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the destination document.
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("This is the main document.");

        // Create the source document to be inserted.
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Font.Color = System.Drawing.Color.Blue;
        srcBuilder.Writeln("This is the inserted document with source formatting.");

        // Insert the source document at the current cursor position, preserving its formatting.
        destBuilder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as HTML.
        string outputPath = "MergedOutput.html";
        destination.Save(outputPath, SaveFormat.Html);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }

        // Optional: simple validation that the file contains content from both documents.
        string htmlContent = File.ReadAllText(outputPath);
        if (!htmlContent.Contains("This is the main document.") || !htmlContent.Contains("This is the inserted document"))
        {
            throw new InvalidOperationException("The merged HTML does not contain expected content from both source documents.");
        }
    }
}
