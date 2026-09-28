using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // Paths for the sample source documents and the merged output.
        string docxPath = Path.Combine(workFolder, "source.docx");
        string rtfPath = Path.Combine(workFolder, "source.rtf");
        string mergedPath = Path.Combine(workFolder, "combined.docx");

        // Create a DOCX source document.
        Document docxSource = new Document();
        DocumentBuilder docxBuilder = new DocumentBuilder(docxSource);
        docxBuilder.Writeln("This is the DOCX source document.");
        docxSource.Save(docxPath, SaveFormat.Docx);

        // Create an RTF source document.
        Document rtfSource = new Document();
        DocumentBuilder rtfBuilder = new DocumentBuilder(rtfSource);
        rtfBuilder.Writeln("This is the RTF source document.");
        rtfSource.Save(rtfPath, SaveFormat.Rtf);

        // Load the destination DOCX document.
        Document destination = new Document(docxPath);

        // Load the RTF document to be appended.
        Document rtfToAppend = new Document(rtfPath);

        // Append the RTF document using destination styles.
        destination.AppendDocument(rtfToAppend, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the merged file exists.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not saved.");

        // Validation: ensure content from both source documents is present.
        string mergedText = destination.GetText();
        if (!mergedText.Contains("This is the DOCX source document.") ||
            !mergedText.Contains("This is the RTF source document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }
    }
}
