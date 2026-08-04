using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the final merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");

        // ---------- Create destination document ----------
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("=== Destination Document Start ===");

        // Protect the destination document with a password.
        const string password = "destPassword";
        destination.Protect(ProtectionType.ReadOnly, password);

        // ---------- Create source document ----------
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("=== Source Document Content ===");

        // Append the source document to the protected destination document.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove protection (password is not required for Unprotect()).
        destination.Unprotect();

        // Save the merged document.
        destination.Save(outputPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged document was not created.");

        // Load the saved document to verify its content.
        Document result = new Document(outputPath);
        string text = result.GetText();

        if (!text.Contains("=== Destination Document Start ===") ||
            !text.Contains("=== Source Document Content ==="))
        {
            throw new InvalidOperationException("The merged document does not contain expected content.");
        }

        // Program ends automatically.
    }
}
