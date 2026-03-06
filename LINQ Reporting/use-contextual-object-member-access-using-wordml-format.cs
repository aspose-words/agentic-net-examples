using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Retrieve the first StructuredDocumentTag (content control) in the document.
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

        // Access the WORDML (Flat OPC) representation of the content control.
        string wordOpenXml = sdt.WordOpenXML;

        // Output the WORDML to the console.
        Console.WriteLine("WORDML of the first StructuredDocumentTag:");
        Console.WriteLine(wordOpenXml);

        // Optionally, modify the document (e.g., add a paragraph).
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Added a new paragraph after accessing WORDML.");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
