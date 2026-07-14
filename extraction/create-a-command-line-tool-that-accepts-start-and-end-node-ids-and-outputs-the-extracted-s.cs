using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExtractSegmentToPdf
{
    public static void Main(string[] args)
    {
        // Prepare a sample document.
        const string samplePath = "sample.docx";
        CreateSampleDocument(samplePath);

        // Parse command‑line arguments (start and end paragraph indices).
        int startIndex = 1;
        int endIndex = 3;
        if (args.Length >= 2 &&
            int.TryParse(args[0], out int s) &&
            int.TryParse(args[1], out int e))
        {
            startIndex = s;
            endIndex = e;
        }

        // Load the source document.
        Document source = new Document(samplePath);
        ParagraphCollection paragraphs = source.FirstSection.Body.Paragraphs;

        // Validate indices.
        if (startIndex < 0 || endIndex < 0 ||
            startIndex >= paragraphs.Count || endIndex >= paragraphs.Count)
        {
            throw new InvalidOperationException("Start or end index is out of range.");
        }

        if (startIndex > endIndex)
        {
            int tmp = startIndex;
            startIndex = endIndex;
            endIndex = tmp;
        }

        // Create a new empty document for the extracted segment.
        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // Import and clone the selected paragraphs into the new document.
        NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = paragraphs[i];
            Node imported = importer.ImportNode(para, true);
            body.AppendChild(imported);
        }

        // Save the extracted segment as PDF.
        const string outputPdf = "extracted.pdf";
        result.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Failed to create the PDF output.");
        }

        // Write a simple log file indicating success.
        File.WriteAllText("extraction.log",
            $"Extracted paragraphs {startIndex}-{endIndex} to {outputPdf}");
    }

    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 0: Introduction.");
        builder.Writeln("Paragraph 1: First content paragraph.");
        builder.Writeln("Paragraph 2: Second content paragraph.");
        builder.Writeln("Paragraph 3: Third content paragraph.");
        builder.Writeln("Paragraph 4: Conclusion.");

        doc.Save(path);
    }
}
