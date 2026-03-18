using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load the binary data that we want to embed.
        // Replace the path with the actual file you wish to package.
        byte[] packageData = File.ReadAllBytes("Path/To/YourFile.zip");

        // Insert the data as an OLE Package (unknown handler) and display it as an icon.
        using (MemoryStream ms = new MemoryStream(packageData))
        {
            // progId "Package" tells Aspose.Words to use the legacy Packager approach.
            Shape oleShape = builder.InsertOleObject(ms, "Package", true, null);

            // Configure the OLE package properties.
            oleShape.OleFormat.OlePackage.FileName = "PackageFile.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Package Display.zip";
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
