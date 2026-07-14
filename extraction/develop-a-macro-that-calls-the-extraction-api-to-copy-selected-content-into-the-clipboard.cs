using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionExample
{
    [STAThread]
    public static void Main()
    {
        // Create a sample document with some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph - plain text.");
        builder.Font.Bold = true;
        builder.Writeln("Second paragraph - bold text.");
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Third paragraph - italic text.");
        doc.Save("sample.docx");

        // Load the document and extract the second paragraph's text.
        Document loaded = new Document("sample.docx");
        if (loaded.FirstSection?.Body?.Paragraphs?.Count < 2)
        {
            throw new InvalidOperationException("The document does not contain the expected second paragraph.");
        }

        Paragraph secondParagraph = loaded.FirstSection.Body.Paragraphs[1];
        if (secondParagraph == null)
        {
            throw new InvalidOperationException("The expected paragraph was not found.");
        }

        string extractedText = secondParagraph.GetText().Trim();

        // Write the extracted text to a file for verification.
        string textFilePath = "extracted.txt";
        File.WriteAllText(textFilePath, extractedText);
        if (!File.Exists(textFilePath))
        {
            throw new InvalidOperationException("Failed to create the extracted text file.");
        }

        // Also create a JSON report containing the extracted content.
        var report = new
        {
            ExtractedParagraphIndex = 2,
            Content = extractedText,
            Timestamp = DateTime.UtcNow
        };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        string jsonFilePath = "extraction_report.json";
        File.WriteAllText(jsonFilePath, json);
        if (!File.Exists(jsonFilePath))
        {
            throw new InvalidOperationException("Failed to create the JSON report file.");
        }

        // Program completed successfully.
    }
}
