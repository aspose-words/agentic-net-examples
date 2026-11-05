using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workDir);

        // Paths for the sample documents.
        string docxPath = Path.Combine(workDir, "source.docx");
        string rtfPath = Path.Combine(workDir, "source.rtf");
        string combinedPath = Path.Combine(workDir, "combined.docx");

        // ---------- Create DOCX source document ----------
        Document docxSource = new Document();
        DocumentBuilder docxBuilder = new DocumentBuilder(docxSource);
        docxBuilder.Writeln("This is the DOCX source document.");
        docxSource.Save(docxPath, SaveFormat.Docx);

        // ---------- Create RTF source document ----------
        Document rtfSource = new Document();
        DocumentBuilder rtfBuilder = new DocumentBuilder(rtfSource);
        rtfBuilder.Writeln("This is the RTF source document.");
        rtfSource.Save(rtfPath, SaveFormat.Rtf);

        // ---------- Load documents for joining ----------
        Document destination = new Document(docxPath);
        Document source = new Document(rtfPath);

        // Append the RTF document to the DOCX using destination styles.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        destination.Save(combinedPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(combinedPath))
        {
            throw new InvalidOperationException("Combined document was not saved.");
        }

        Document combined = new Document(combinedPath);
        string combinedText = combined.GetText();

        if (!combinedText.Contains("This is the DOCX source document.") ||
            !combinedText.Contains("This is the RTF source document."))
        {
            throw new InvalidOperationException("Combined document does not contain expected content from both source files.");
        }

        // Clean up temporary files (optional).
        // File.Delete(docxPath);
        // File.Delete(rtfPath);
        // File.Delete(combinedPath);
        // Directory.Delete(workDir);
    }
}
