using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionExample
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        builder.Writeln("Paragraph 5");
        sourceDoc.Save("sample.docx");

        // Load the document for extraction.
        Document loadedDoc = new Document("sample.docx");
        Body body = loadedDoc.FirstSection.Body;

        // Define start and end paragraphs (intentionally reversed to trigger error handling).
        Paragraph startParagraph = body.Paragraphs[3]; // "Paragraph 4"
        Paragraph endParagraph = body.Paragraphs[1];   // "Paragraph 2"

        try
        {
            // Validate that start node appears before end node.
            int startIndex = body.Paragraphs.IndexOf(startParagraph);
            int endIndex = body.Paragraphs.IndexOf(endParagraph);
            if (startIndex < 0 || endIndex < 0)
                throw new InvalidOperationException("One of the specified paragraphs was not found.");

            if (startIndex > endIndex)
                throw new InvalidOperationException("Start paragraph appears after the end paragraph. Extraction aborted.");

            // Extract the range of paragraphs into a new document.
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren();

            Section resultSection = new Section(resultDoc);
            resultDoc.AppendChild(resultSection);
            Body resultBody = new Body(resultDoc);
            resultSection.AppendChild(resultBody);

            for (int i = startIndex; i <= endIndex; i++)
            {
                Paragraph para = (Paragraph)body.Paragraphs[i].Clone(true);
                resultBody.AppendChild(para);
            }

            // Save the extracted content.
            string outputPath = "extracted.docx";
            resultDoc.Save(outputPath);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Extraction succeeded but the output file was not created.");

            Console.WriteLine("Extraction completed successfully.");
        }
        catch (Exception ex)
        {
            // Handle errors such as reversed node order.
            Console.WriteLine($"Error during extraction: {ex.Message}");
        }
    }
}
