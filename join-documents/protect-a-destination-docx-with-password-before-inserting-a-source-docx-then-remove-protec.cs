using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string destinationPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string mergedPath = "Merged.docx";

        // Create a sample destination document.
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // Create a sample source document.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the destination document and protect it with a password.
        Document protectedDest = new Document(destinationPath);
        protectedDest.Protect(ProtectionType.ReadOnly, "myPassword");

        // Load the source document.
        Document source = new Document(sourcePath);

        // Append the source document to the protected destination.
        protectedDest.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove protection from the merged document.
        protectedDest.Unprotect();

        // Save the final merged document.
        protectedDest.Save(mergedPath, SaveFormat.Docx);

        // Validate that the merged file was created.
        if (!File.Exists(mergedPath))
        {
            throw new InvalidOperationException("The merged document was not saved correctly.");
        }

        // Optional validation: ensure the merged document contains content from both sources.
        Document result = new Document(mergedPath);
        int paragraphCount = result.GetChildNodes(NodeType.Paragraph, true).Count;
        if (paragraphCount < 2)
        {
            throw new InvalidOperationException("The merged document does not contain expected content.");
        }
    }
}
