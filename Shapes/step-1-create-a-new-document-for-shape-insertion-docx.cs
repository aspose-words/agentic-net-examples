using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ShapeInsertionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document (optional for further shape insertion).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure OOXML save options to use a compliance level that supports DML shapes.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document as a DOCX file.
        doc.Save("ShapeInsertion.docx", saveOptions);
    }
}
