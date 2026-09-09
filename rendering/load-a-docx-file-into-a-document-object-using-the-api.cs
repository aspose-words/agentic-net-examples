using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the sample DOCX file.
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");

        // Create a simple document and save it.
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);
        builder.Writeln("Hello World!");
        createDoc.Save(samplePath);

        // Load the DOCX file into a new Document object.
        Document loadedDoc = new Document(samplePath);

        // Verify loading by printing the document text.
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}
