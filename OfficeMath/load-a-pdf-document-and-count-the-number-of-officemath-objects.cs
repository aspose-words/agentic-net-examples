using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the PDF file to be processed.
        string pdfPath = "input.pdf";

        // Load options for PDF. Enable conversion of EquationXML shapes to OfficeMath objects.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;

        // Load the PDF document using the provided Document constructor that accepts a file name and LoadOptions.
        Document doc = new Document(pdfPath, loadOptions);

        // Retrieve all OfficeMath nodes in the document (search recursively).
        var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Count the OfficeMath objects.
        int officeMathCount = officeMathNodes.Count;

        // Output the result.
        Console.WriteLine($"Number of OfficeMath objects: {officeMathCount}");
    }
}
