using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents
        CreateSampleDocuments(inputDir);

        // Get list of input files
        string[] inputFiles = Directory.GetFiles(inputDir, "*.docx");

        // List to hold report entries
        List<ExtractionReport> reports = new List<ExtractionReport>();

        // Parallel extraction
        Parallel.ForEach(inputFiles, inputFile =>
        {
            // Load source document
            Document sourceDoc = new Document(inputFile);

            // Validate that the document has at least one paragraph
            Paragraph firstParagraph = sourceDoc.FirstSection?.Body?.Paragraphs?[0];
            if (firstParagraph == null)
                throw new InvalidOperationException($"Document '{inputFile}' does not contain a paragraph to extract.");

            // Create a new document to hold the extracted range
            Document extractedDoc = new Document();
            extractedDoc.RemoveAllChildren();

            // Build minimal document structure
            Section section = new Section(extractedDoc);
            extractedDoc.AppendChild(section);
            Body body = new Body(extractedDoc);
            section.AppendChild(body);

            // Import the paragraph from the source document into the new document
            Node importedParagraph = extractedDoc.ImportNode(firstParagraph, true);
            body.AppendChild(importedParagraph);

            // Determine output file names
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputFile);
            string extractedDocPath = Path.Combine(outputDir, fileNameWithoutExt + "_Extracted.docx");
            string extractedTextPath = Path.Combine(outputDir, fileNameWithoutExt + "_Extracted.txt");

            // Save the extracted document
            extractedDoc.Save(extractedDocPath);

            // Save the extracted text to a plain text file
            File.WriteAllText(extractedTextPath, firstParagraph.GetText());

            // Add entry to the report
            lock (reports)
            {
                reports.Add(new ExtractionReport
                {
                    SourceDocument = inputFile,
                    ExtractedDocument = extractedDocPath,
                    ExtractedTextFile = extractedTextPath,
                    ExtractedTextLength = firstParagraph.GetText().Length
                });
            }
        });

        // Validate that all expected output files were created
        foreach (ExtractionReport report in reports)
        {
            if (!File.Exists(report.ExtractedDocument))
                throw new InvalidOperationException($"Expected extracted document was not created: {report.ExtractedDocument}");
            if (!File.Exists(report.ExtractedTextFile))
                throw new InvalidOperationException($"Expected extracted text file was not created: {report.ExtractedTextFile}");
        }

        // Serialize report to JSON
        string reportJsonPath = Path.Combine(outputDir, "ExtractionReport.json");
        string json = JsonConvert.SerializeObject(reports, Formatting.Indented);
        File.WriteAllText(reportJsonPath, json);

        // Final validation
        if (!File.Exists(reportJsonPath))
            throw new InvalidOperationException("Extraction report JSON file was not created.");
    }

    private static void CreateSampleDocuments(string folderPath)
    {
        for (int i = 1; i <= 5; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} Title");
            builder.Writeln($"This is the body content of document {i}. It contains sample text for extraction.");
            string filePath = Path.Combine(folderPath, $"Doc{i}.docx");
            doc.Save(filePath);
        }
    }

    private class ExtractionReport
    {
        public string SourceDocument { get; set; } = string.Empty;
        public string ExtractedDocument { get; set; } = string.Empty;
        public string ExtractedTextFile { get; set; } = string.Empty;
        public int ExtractedTextLength { get; set; }
    }
}
