using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // File paths
        string docxPath = Path.Combine(outputDir, "SourceDocument.docx");
        string rtfPath = Path.Combine(outputDir, "SourceDocument.rtf");
        string combinedPath = Path.Combine(outputDir, "CombinedDocument.docx");

        // Create a DOCX source document
        var docxSource = new Document();
        var builderDocx = new DocumentBuilder(docxSource);
        builderDocx.Writeln("This is the DOCX source document.");
        docxSource.Save(docxPath, SaveFormat.Docx);

        // Create an RTF source document
        var rtfSource = new Document();
        var builderRtf = new DocumentBuilder(rtfSource);
        builderRtf.Writeln("This is the RTF source document.");
        rtfSource.Save(rtfPath, SaveFormat.Rtf);

        // Load the destination DOCX and the RTF to be appended
        var destination = new Document(docxPath);
        var rtfToAppend = new Document(rtfPath);

        // Append the RTF document using destination styles
        destination.AppendDocument(rtfToAppend, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX
        destination.Save(combinedPath, SaveFormat.Docx);

        // Validation: ensure the file exists
        if (!File.Exists(combinedPath))
            throw new InvalidOperationException("The combined document was not saved.");

        // Validation: ensure both source texts are present
        var combinedDoc = new Document(combinedPath);
        string combinedText = combinedDoc.GetText();

        if (!combinedText.Contains("This is the DOCX source document.") ||
            !combinedText.Contains("This is the RTF source document."))
        {
            throw new InvalidOperationException("The combined document does not contain expected content.");
        }

        // Execution completed successfully
    }
}
