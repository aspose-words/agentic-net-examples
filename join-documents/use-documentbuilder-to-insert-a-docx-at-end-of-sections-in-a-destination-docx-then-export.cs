using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string folderPath = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(folderPath);

        // Define file paths for the destination, source, and merged output documents.
        string destPath = Path.Combine(folderPath, "Destination.docx");
        string sourcePath = Path.Combine(folderPath, "Source.docx");
        string outputPath = Path.Combine(folderPath, "MergedOutput.docx");

        // ------------------------------------------------------------
        // Create a destination document with two sections.
        // ------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document - Section 1");
        destBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        destBuilder.Writeln("Destination Document - Section 2");
        destDoc.Save(destPath, SaveFormat.Docx);

        // ------------------------------------------------------------
        // Create a source document that will be inserted.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("Inserted Source Document Content");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ------------------------------------------------------------
        // Load the documents for joining.
        // ------------------------------------------------------------
        Document destination = new Document(destPath);
        Document source = new Document(sourcePath);

        // Insert the source document after the last paragraph of each section in the destination.
        for (int i = 0; i < destination.Sections.Count; i++)
        {
            Section currentSection = destination.Sections[i];
            // Get the last paragraph of the current section.
            Paragraph lastParagraph = currentSection.Body.Paragraphs[currentSection.Body.Paragraphs.Count - 1];

            // Position the builder at that paragraph.
            DocumentBuilder builder = new DocumentBuilder(destination);
            builder.MoveTo(lastParagraph);

            // Insert the source document after the current position.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as DOCX.
        destination.Save(outputPath, SaveFormat.Docx);

        // ------------------------------------------------------------
        // Validation: ensure the output file exists and contains source content.
        // ------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new Exception("Merged output file was not created.");

        Document merged = new Document(outputPath);
        if (!merged.Range.Text.Contains("Inserted Source Document Content"))
            throw new Exception("Source content was not found in the merged document.");

        // Optional cleanup (commented out for inspection).
        // File.Delete(destPath);
        // File.Delete(sourcePath);
        // File.Delete(outputPath);
        // Directory.Delete(folderPath);
    }
}
