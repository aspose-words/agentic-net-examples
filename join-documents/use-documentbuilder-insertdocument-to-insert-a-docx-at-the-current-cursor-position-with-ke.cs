using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string destPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string outputPath = "MergedOutput.html";

        // Create destination document with initial content
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the beginning of the destination document.");

        // Save destination document (optional, just to have a file)
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create source document that will be inserted
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This paragraph comes from the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the source document from the saved DOCX file
        Document sourceToInsert = new Document(sourcePath);

        // Insert the source document at the current cursor position with KeepSourceFormatting
        destBuilder.Writeln("Content before insertion.");
        destBuilder.InsertDocument(sourceToInsert, ImportFormatMode.KeepSourceFormatting);
        destBuilder.Writeln("Content after insertion.");

        // Save the merged document as HTML
        destDoc.Save(outputPath, SaveFormat.Html);

        // Validate that the output file was created
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }

        // Clean up temporary files (optional)
        File.Delete(destPath);
        File.Delete(sourcePath);
    }
}
