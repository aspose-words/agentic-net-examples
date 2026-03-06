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

        // Load options: enable conversion of shapes that contain EquationXML into OfficeMath objects.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;

        // Load the PDF document using the provided Document constructor that accepts a file name and LoadOptions.
        Document doc = new Document(pdfPath, loadOptions);

        // Retrieve all OfficeMath nodes in the document (including those nested inside other nodes).
        int officeMathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;

        // Output the count.
        Console.WriteLine($"Number of OfficeMath objects: {officeMathCount}");
    }
}
