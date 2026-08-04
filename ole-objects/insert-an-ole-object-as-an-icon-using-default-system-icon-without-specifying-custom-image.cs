using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple text file that will be embedded as an OLE object.
        string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.txt");
        File.WriteAllText(tempFilePath, "This is a sample text file for OLE embedding.");

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the text file as an OLE object displayed as an icon.
        // Pass null for iconFile and iconCaption to use the default system icon and file name as caption.
        builder.InsertOleObjectAsIcon(tempFilePath, false, null, null);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectIcon.docx");
        doc.Save(outputPath);
    }
}
