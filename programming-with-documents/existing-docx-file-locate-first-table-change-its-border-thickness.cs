using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Change the thickness of all four outline borders.
        // The last parameter (true) removes any explicit cell borders.
        table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
