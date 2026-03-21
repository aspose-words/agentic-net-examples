using System;
using Aspose.Words;

class InsertExternalHtml
{
    static void Main()
    {
        // Sample HTML content to insert.
        string htmlContent = "<p>Hello <b>World</b> from <i>embedded</i> HTML snippet.</p>";

        // Path where the resulting Word document will be saved (in the current directory).
        string outputDocPath = System.IO.Path.Combine(Environment.CurrentDirectory, "Result.docx");

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the HTML content into the document at the current cursor position.
        builder.InsertHtml(htmlContent);

        // Save the document to the specified location.
        doc.Save(outputDocPath);

        Console.WriteLine($"Document saved to: {outputDocPath}");
    }
}
