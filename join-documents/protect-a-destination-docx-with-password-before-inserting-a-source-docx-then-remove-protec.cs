using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string destPath = "Destination.docx";
        const string srcPath = "Source.docx";
        const string outputPath = "MergedOutput.docx";
        const string password = "Secret123";

        // Create destination document with sample content
        var destinationDoc = new Document();
        var destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");
        destinationDoc.Save(destPath, SaveFormat.Docx);

        // Protect the destination document with a password
        destinationDoc.Protect(ProtectionType.ReadOnly, password);
        // Save the protected version (overwrites the previous file)
        destinationDoc.Save(destPath, SaveFormat.Docx);

        // Create source document with sample content
        var sourceDoc = new Document();
        var srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        sourceDoc.Save(srcPath, SaveFormat.Docx);

        // Load the protected destination document
        var protectedDest = new Document(destPath);
        // Load the source document
        var source = new Document(srcPath);

        // Append the source document to the protected destination
        protectedDest.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove protection before saving the final merged document
        protectedDest.Unprotect();

        // Save the merged, unprotected document
        protectedDest.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file exists
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Merged output file was not created.");

        // Optional validation: ensure content from both documents is present
        var merged = new Document(outputPath);
        string mergedText = merged.GetText();
        if (!mergedText.Contains("destination document") || !mergedText.Contains("source document"))
            throw new InvalidOperationException("Merged document does not contain expected content.");
    }
}
