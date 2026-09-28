using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        string sourcePath = "Sample.docx";

        // Create a simple DOCX file if it does not already exist.
        if (!File.Exists(sourcePath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words!");
            doc.Save(sourcePath);
        }

        // Load the DOCX file into a Document object.
        Document loadedDoc = new Document(sourcePath);

        // Validate that the document was loaded correctly.
        if (loadedDoc.PageCount < 1)
        {
            throw new InvalidOperationException("Loaded document contains no pages.");
        }

        // Save a copy to confirm that loading succeeded.
        string copyPath = "LoadedCopy.docx";
        loadedDoc.Save(copyPath);

        // Verify the copy was created.
        if (!File.Exists(copyPath))
        {
            throw new FileNotFoundException("Failed to save the loaded document copy.", copyPath);
        }

        // Indicate successful execution.
        Console.WriteLine("Document loaded and saved successfully.");
    }
}
