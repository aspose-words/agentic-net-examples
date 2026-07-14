using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple DOC document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");

        // Save the document to a local DOC file (bootstrap step required by the rules).
        const string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Simulate a network response using a MemoryStream.
        using (MemoryStream responseStream = new MemoryStream())
        {
            // Convert and write the PDF directly to the simulated response stream.
            doc.Save(responseStream, SaveFormat.Pdf);

            // Verify that data was written.
            if (responseStream.Length == 0)
                throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

            // (Optional) Reset position if further reading is needed.
            responseStream.Position = 0;
        }

        // Clean up the temporary DOC file.
        if (File.Exists(inputPath))
            File.Delete(inputPath);
    }
}
