using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractContentBetweenControls
{
    static void Main()
    {
        // Load the RTF document.
        Document doc = new Document("InputDocument.rtf");

        // Define the titles (or tags) of the start and end content controls.
        // Adjust these values to match the actual content controls in your document.
        const string startControlTitle = "StartControl";
        const string endControlTitle = "EndControl";

        // Find the start and end StructuredDocumentTagRangeStart nodes.
        StructuredDocumentTagRangeStart startTag = null;
        StructuredDocumentTagRangeEnd endTag = null;

        // Search all nodes for the start and end tags.
        foreach (Node node in doc.GetChildNodes(NodeType.StructuredDocumentTagRangeStart, true))
        {
            var sdtStart = (StructuredDocumentTagRangeStart)node;
            if (sdtStart.Title == startControlTitle || sdtStart.Tag == startControlTitle)
            {
                startTag = sdtStart;
                break;
            }
        }

        foreach (Node node in doc.GetChildNodes(NodeType.StructuredDocumentTagRangeEnd, true))
        {
            var sdtEnd = (StructuredDocumentTagRangeEnd)node;
            // The end tag shares the same Id as its start tag.
            if (startTag != null && sdtEnd.Id == startTag.Id)
            {
                endTag = sdtEnd;
                break;
            }
        }

        if (startTag == null || endTag == null)
        {
            Console.WriteLine("Could not locate the specified content controls.");
            return;
        }

        // The range between the start and end tags is represented by the start tag's Range.
        // It includes all nodes up to (but not including) the end tag.
        string extractedText = startTag.Range.Text;

        // Optionally, write the extracted text to a separate file.
        File.WriteAllText("ExtractedContent.txt", extractedText);

        // If you need the extracted portion as a separate Word document:
        Document extractedDoc = startTag.Range.ToDocument();
        extractedDoc.Save("ExtractedContent.docx");

        Console.WriteLine("Extraction completed successfully.");
    }
}
