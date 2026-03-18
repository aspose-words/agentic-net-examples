using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Apply a text watermark to every page of the document.
        doc.Watermark.SetText("Confidential");

        // Save the document in DOCX format.
        doc.Save("WatermarkedDocument.docx");
    }
}
