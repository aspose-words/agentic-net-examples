using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary DOC file path.
        string docPath = Path.Combine(Path.GetTempPath(), "Sample.doc");

        // Build a simple DOC document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Save the document in DOC format (binary Word 97‑2003).
        doc.Save(docPath, SaveFormat.Doc);

        // Load the saved DOC file.
        Document loadedDoc = new Document(docPath);

        // Simulate a network response using a MemoryStream.
        using (MemoryStream responseStream = new MemoryStream())
        {
            // Convert the document to PDF directly into the stream.
            loadedDoc.Save(responseStream, SaveFormat.Pdf);

            // Validate that the stream contains data.
            if (responseStream.Length == 0)
                throw new InvalidOperationException("The PDF stream is empty.");

            // Reset the stream position for any further reads.
            responseStream.Position = 0;

            // Example output: display the size of the generated PDF.
            Console.WriteLine($"PDF stream length: {responseStream.Length} bytes");
        }

        // Clean up the temporary DOC file.
        if (File.Exists(docPath))
            File.Delete(docPath);
    }
}
