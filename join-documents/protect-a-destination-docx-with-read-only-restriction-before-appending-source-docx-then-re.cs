using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string destPath = "dest.docx";
        const string sourcePath = "source.docx";
        const string outputPdfPath = "merged.pdf";
        const string protectionPassword = "pwd";

        // -----------------------------------------------------------------
        // Create destination document with sample content
        // -----------------------------------------------------------------
        var destDoc = new Document();
        var builder = new DocumentBuilder(destDoc);
        builder.Writeln("This is the destination document.");
        destDoc.Save(destPath);

        // Protect the destination document with read‑only restriction
        destDoc.Protect(ProtectionType.ReadOnly, protectionPassword);
        destDoc.Save(destPath); // Overwrite with protection applied

        // -----------------------------------------------------------------
        // Create source document with sample content
        // -----------------------------------------------------------------
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the protected destination document and the source document
        // -----------------------------------------------------------------
        var protectedDest = new Document(destPath);
        var source = new Document(sourcePath);

        // Append source document to the protected destination document
        protectedDest.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove protection after appending
        protectedDest.Unprotect(protectionPassword);

        // Save the merged document as PDF
        protectedDest.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new Exception("The merged PDF file was not created.");
        }

        // Ensure the merged document contains content from both source documents
        int paragraphCount = protectedDest.GetChildNodes(NodeType.Paragraph, true).Count;
        if (paragraphCount < 2)
        {
            throw new Exception("The merged document does not contain expected content from both source documents.");
        }
    }
}
