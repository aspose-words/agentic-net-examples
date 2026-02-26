using System;
using Aspose.Words;

class ExtractRangeText
{
    static void Main()
    {
        // Load the source DOCX document (load rule).
        string inputFile = "input.docx";
        Document sourceDoc = new Document(inputFile);

        // Extract the full text from the document's range.
        string extractedText = sourceDoc.Range.Text;

        // Create a new blank document (create rule) to hold the extracted text.
        Document outputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(outputDoc);

        // Write the extracted text into the new document.
        builder.Writeln(extractedText);

        // Save the result as a plain‑text file (save rule).
        string outputFile = "extracted.txt";
        outputDoc.Save(outputFile, SaveFormat.Text);
    }
}
