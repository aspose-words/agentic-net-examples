using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlBindingExample
{
    static void Main()
    {
        // 1. Create a new blank document (DOCM will be set on save).
        Document doc = new Document();

        // 2. Add a custom XML part that will hold the data to bind.
        //    Use a GUID as the part identifier.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent =
            "<root>" +
                "<person>" +
                    "<name>John Doe</name>" +
                    "<email>john.doe@example.com</email>" +
                "</person>" +
            "</root>";

        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // 3. Create a content control (structured document tag) that will display the name.
        //    Use a plain‑text content control placed at block level.
        StructuredDocumentTag nameControl = new StructuredDocumentTag(
            doc, SdtType.PlainText, MarkupLevel.Block);

        // 4. Bind the content control to the <name> element of the custom XML part.
        //    The XPath points to the first <name> element under the first <person>.
        nameControl.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/name[1]", string.Empty);

        // 5. Insert the content control into the document body.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Employee Name:");
        builder.InsertNode(nameControl);
        builder.Writeln(); // optional line break after the control

        // 6. Save the document as a macro‑enabled Word file (.docm).
        //    The SaveFormat ensures the correct file type regardless of extension.
        doc.Save("ContentControlBound.docm", SaveFormat.Docm);
    }
}
