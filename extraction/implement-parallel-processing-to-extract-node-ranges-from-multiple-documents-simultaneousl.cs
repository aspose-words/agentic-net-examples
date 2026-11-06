using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;

public class ParallelExtractionExample
{
    public static void Main()
    {
        // Prepare input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a set of sample documents.
        int documentCount = 5;
        for (int i = 0; i < documentCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} - Intro paragraph.");
            builder.Writeln($"Document {i} - Target paragraph 1.");
            builder.Writeln($"Document {i} - Target paragraph 2.");
            builder.Writeln($"Document {i} - End paragraph.");
            string filePath = Path.Combine(inputDir, $"doc{i}.docx");
            doc.Save(filePath);
        }

        // Gather all input files.
        string[] inputFiles = Directory.GetFiles(inputDir, "*.docx");

        // Process each document in parallel, extracting a specific range of paragraphs.
        Parallel.ForEach(inputFiles, inputFile =>
        {
            // Load the source document.
            Document source = new Document(inputFile);

            // Identify the paragraphs to extract (indexes 1 and 2).
            Paragraph startParagraph = source.FirstSection?.Body?.Paragraphs[1];
            Paragraph endParagraph = source.FirstSection?.Body?.Paragraphs[2];
            if (startParagraph == null || endParagraph == null)
                throw new InvalidOperationException($"Required paragraphs not found in {Path.GetFileName(inputFile)}.");

            // Create a new empty document that will hold the extracted range.
            Document extracted = new Document();
            extracted.RemoveAllChildren(); // Remove the default empty section.
            Section section = new Section(extracted);
            extracted.AppendChild(section);
            Body body = new Body(extracted);
            section.AppendChild(body);

            // Use a NodeImporter to import nodes from the source document into the destination document.
            NodeImporter importer = new NodeImporter(source, extracted, ImportFormatMode.KeepSourceFormatting);
            Node importedStart = importer.ImportNode(startParagraph, true);
            Node importedEnd = importer.ImportNode(endParagraph, true);

            // Append the imported paragraphs to the new document's body.
            body.AppendChild(importedStart);
            body.AppendChild(importedEnd);

            // Define output file paths.
            string baseName = Path.GetFileNameWithoutExtension(inputFile);
            string outputDocPath = Path.Combine(outputDir, $"{baseName}_extracted.docx");
            string outputTxtPath = Path.Combine(outputDir, $"{baseName}_extracted.txt");

            // Save the extracted document.
            extracted.Save(outputDocPath);

            // Also write the extracted plain text to a .txt file.
            string extractedText = extracted.GetText();
            File.WriteAllText(outputTxtPath, extractedText);

            // Verify that the output files were created.
            if (!File.Exists(outputDocPath) || !File.Exists(outputTxtPath))
                throw new InvalidOperationException($"Extraction output not created for {baseName}.");
        });

        // Extraction completed.
    }
}
