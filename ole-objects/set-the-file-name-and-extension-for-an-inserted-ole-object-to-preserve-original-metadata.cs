using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Dummy data to simulate a file (e.g., a ZIP archive) that will be embedded as an OLE package.
        byte[] dummyData = Encoding.UTF8.GetBytes("Dummy content for OLE package");
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // Insert the OLE object as a package and display it as an icon.
            Shape shape = builder.InsertOleObject(stream, "Package", true, null);

            // Preserve original metadata by setting the file name and display name.
            shape.OleFormat.OlePackage.FileName = "OriginalFileName.zip";
            shape.OleFormat.OlePackage.DisplayName = "OriginalFileName.zip";
        }

        // Ensure the output directory exists.
        string outputDir = "Artifacts";
        Directory.CreateDirectory(outputDir);

        // Save the document to a file.
        doc.Save(Path.Combine(outputDir, "OlePackageExample.docx"));
    }
}
