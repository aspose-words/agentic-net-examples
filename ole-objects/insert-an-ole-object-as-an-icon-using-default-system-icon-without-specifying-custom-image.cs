using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a sample file to embed as an OLE object.
        string oleFilePath = "sample.txt";
        File.WriteAllText(oleFilePath, "This is a sample text file embedded as an OLE object.");

        // Insert the OLE object as an icon using the default system icon.
        // Use the "Package" progId to embed any file type.
        using (FileStream stream = File.OpenRead(oleFilePath))
        {
            // progId = "Package", asIcon = true, iconFileName = null (default system icon)
            builder.InsertOleObject(stream, "Package", true, null);
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
