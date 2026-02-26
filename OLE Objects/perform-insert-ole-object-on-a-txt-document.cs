using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoTxt
{
    static void Main()
    {
        // Load a plain‑text file into a Word document.
        Document doc = new Document("input.txt");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a caption before the OLE object.
        builder.Writeln("Embedded OLE object:");

        // Open the file that will be embedded (e.g., a ZIP archive) as a stream.
        using (FileStream oleStream = File.Open("sample.zip", FileMode.Open))
        {
            // Optional: open an icon image to represent the OLE object.
            using (FileStream iconStream = File.Open("icon.ico", FileMode.Open))
            {
                // Insert the OLE object as an icon.
                // progId "Package" is used for generic packages.
                // asIcon = true to display the object as an icon.
                builder.InsertOleObject(oleStream, "Package", true, iconStream);
            }
        }

        // Save the resulting document (Word format) to disk.
        doc.Save("output.docx");
    }
}
