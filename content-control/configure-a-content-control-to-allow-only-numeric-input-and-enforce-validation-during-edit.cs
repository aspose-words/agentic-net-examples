using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that explains the purpose of the content control.
        builder.Writeln("Enter a numeric value:");

        // Create an inline plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NumericInput",
            Tag = "NumericInputTag"
        };

        // Insert the content control into the document.
        builder.InsertNode(numericSdt);

        // Create a custom XML part that will hold the bound data.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><num>0</num></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // (Optional) Register a namespace for the XML part. Word uses the namespace for validation.
        xmlPart.Schemas.Add("http://example.com");

        // Bind the content control to the <num> element inside the custom XML part.
        // Word will enforce the data type defined by the XML schema (integer) during editing.
        numericSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/num[1]", string.Empty);

        // Save the resulting document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NumericContentControl.docx");
        doc.Save(outputPath);
    }
}
