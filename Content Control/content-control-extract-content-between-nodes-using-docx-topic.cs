using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOCX file from disk.
        Document doc = new Document("Input.docx");

        // Locate the StructuredDocumentTag (content control) with a specific tag.
        StructuredDocumentTag targetTag = null;
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Replace "MyContentControl" with the actual Tag value of the control you need.
            if (sdt.Tag == "MyContentControl")
            {
                targetTag = sdt;
                break;
            }
        }

        if (targetTag != null)
        {
            // Extract the text that resides inside the content control.
            string extractedText = targetTag.Range.Text.Trim();

            Console.WriteLine("Extracted text: " + extractedText);

            // Create a new document to hold the extracted content.
            Document extractedDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(extractedDoc);
            builder.Writeln(extractedText);

            // Save the new document containing only the extracted content.
            extractedDoc.Save("Extracted.docx");
        }
        else
        {
            Console.WriteLine("Content control with the specified tag was not found.");
        }
    }
}
