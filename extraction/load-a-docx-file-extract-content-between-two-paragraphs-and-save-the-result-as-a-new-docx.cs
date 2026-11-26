using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample source document with several paragraphs.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Paragraph before start");
        builder.Writeln("Start paragraph");          // Start boundary
        builder.Writeln("Middle paragraph 1");
        builder.Writeln("Middle paragraph 2");
        builder.Writeln("End paragraph");            // End boundary
        builder.Writeln("Paragraph after end");
        source.Save("source.docx");

        // -----------------------------------------------------------------
        // Load the source document.
        // -----------------------------------------------------------------
        Document loaded = new Document("source.docx");

        // Locate the start and end paragraphs (by index).
        Body body = loaded.FirstSection.Body;
        int startIndex = 1; // "Start paragraph"
        int endIndex = 4;   // "End paragraph"

        if (startIndex < 0 || endIndex >= body.Paragraphs.Count || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries.");

        // -----------------------------------------------------------------
        // Prepare a fresh destination document.
        // -----------------------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section resultSection = new Section(result);
        result.AppendChild(resultSection);

        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // Import and append each paragraph from start to end inclusive.
        // -----------------------------------------------------------------
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = body.Paragraphs[i];
            if (para == null)
                throw new InvalidOperationException($"Paragraph at index {i} not found.");

            // Import the paragraph into the destination document.
            Node imported = result.ImportNode(para, true);
            resultBody.AppendChild(imported);
        }

        // -----------------------------------------------------------------
        // Validate the extracted paragraph count.
        // -----------------------------------------------------------------
        int expectedCount = endIndex - startIndex + 1;
        if (result.FirstSection.Body.Paragraphs.Count != expectedCount)
            throw new InvalidOperationException("Extracted paragraph count does not match expected.");

        // -----------------------------------------------------------------
        // Save the extracted content as a new DOCX file.
        // -----------------------------------------------------------------
        string outputPath = "extracted.docx";
        result.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
