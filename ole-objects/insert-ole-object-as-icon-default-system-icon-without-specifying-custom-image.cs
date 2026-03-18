using Aspose.Words;
using System;

class InsertOleIconExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        // Replace with the actual path to your file.
        string oleFilePath = @"C:\Temp\Sample.xlsx";

        // Insert the OLE object as an icon.
        // - isLinked = false (embed the file).
        // - iconFile = null (use the default system icon for the file type).
        // - iconCaption = null (use the file name as the caption).
        builder.InsertOleObjectAsIcon(oleFilePath, false, null, null);

        // Save the resulting document.
        doc.Save(@"C:\Temp\OleIcon.docx");
    }
}
