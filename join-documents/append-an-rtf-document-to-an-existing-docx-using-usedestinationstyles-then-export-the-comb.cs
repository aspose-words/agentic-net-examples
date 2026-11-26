using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the destination DOCX document.
        string destPath = "Destination.docx";
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        destBuilder.Writeln("Destination Document Heading");
        destBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        destBuilder.Writeln("This is the body of the destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create the source RTF document.
        string sourcePath = "Source.rtf";
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Source Document Heading");
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("This is the body of the source RTF document.");
        sourceDoc.Save(sourcePath, SaveFormat.Rtf);

        // Load both documents.
        Document destination = new Document(destPath);
        Document source = new Document(sourcePath);

        // Append the RTF document to the DOCX using destination styles.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        string combinedPath = "Combined.docx";
        destination.Save(combinedPath, SaveFormat.Docx);

        // Validate that the combined file exists and contains content from both sources.
        if (!File.Exists(combinedPath))
            throw new InvalidOperationException("Combined document was not saved.");

        Document combined = new Document(combinedPath);
        if (combined.Sections.Count < 2)
            throw new InvalidOperationException("Combined document does not contain the expected sections.");

        string combinedText = combined.GetText();
        if (!combinedText.Contains("Destination Document Heading") ||
            !combinedText.Contains("Source Document Heading"))
            throw new InvalidOperationException("Combined document is missing expected content.");
    }
}
