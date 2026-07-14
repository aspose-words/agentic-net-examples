using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path of the sample DOCX file.
        string sourcePath = Path.Combine(artifactsDir, "Sample.docx");

        // Create a sample DOCX document if it does not already exist.
        if (!File.Exists(sourcePath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words!");
            doc.Save(sourcePath, SaveFormat.Docx);
        }

        // Load the DOCX file into a Document object.
        Document loadedDoc = new Document(sourcePath);

        // Verify that the document was loaded correctly.
        string loadedText = loadedDoc.GetText();
        if (!loadedText.Contains("Hello Aspose.Words!"))
        {
            throw new InvalidOperationException("Loaded document does not contain the expected text.");
        }

        // Save a copy to demonstrate that the loaded document can be saved again.
        string copyPath = Path.Combine(artifactsDir, "Copy.docx");
        loadedDoc.Save(copyPath, SaveFormat.Docx);
    }
}
