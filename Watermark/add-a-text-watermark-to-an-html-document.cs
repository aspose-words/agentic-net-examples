using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");

        // Add a text watermark to every page of the document.
        // The watermark text can be customized as needed.
        doc.Watermark.SetText("Confidential");

        // Save the document back to HTML (or any other supported format).
        doc.Save("output.html");
    }
}
