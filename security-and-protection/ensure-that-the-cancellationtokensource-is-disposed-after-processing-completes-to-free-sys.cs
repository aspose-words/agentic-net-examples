using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Use CancellationTokenSource to represent a cancellable operation.
        // The using statement guarantees disposal after the operation completes.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Perform the document save operation.
            // Aspose.Words does not require the token for Save, but the token source
            // is kept to illustrate proper disposal after processing.
            doc.Save("output.docx");
            // No cancellation is requested; the operation finishes normally.
        } // cts is disposed here, freeing system resources.

        // Validate that the output file was created.
        if (!File.Exists("output.docx"))
            throw new Exception("The document was not saved as expected.");

        // Clean up the generated file.
        File.Delete("output.docx");
    }
}
