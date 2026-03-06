using System;
using Aspose.Words;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Recalculate word count statistics and include line count.
        // The overload with 'true' updates the Lines property.
        doc.UpdateWordCount(true);

        // Retrieve the estimated number of lines in the document.
        int lineCount = doc.BuiltInDocumentProperties.Lines;

        // Optionally, also retrieve the number of paragraphs.
        int paragraphCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;

        // Output the results.
        Console.WriteLine($"Lines: {lineCount}");
        Console.WriteLine($"Paragraphs: {paragraphCount}");

        // Save the document if any modifications were made (optional).
        doc.Save("Output.docx");
    }
}
