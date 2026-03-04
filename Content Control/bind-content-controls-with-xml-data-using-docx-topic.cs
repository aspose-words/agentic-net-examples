using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlXmlBinding
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define XML data that will be stored in a custom XML part.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"
            <catalog>
                <book>
                    <title>Everyday Italian</title>
                    <author>Giada De Laurentiis</author>
                </book>
                <book>
                    <title>The C Programming Language</title>
                    <author>Brian W. Kernighan, Dennis M. Ritchie</author>
                </book>
            </catalog>";

        // Add the custom XML part to the document.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Create a content control (structured document tag) that will display the title of the first book.
        StructuredDocumentTag titleTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        // Map the content control to the XML node using XPath.
        titleTag.XmlMapping.SetMapping(xmlPart, "/catalog[1]/book[1]/title[1]", string.Empty);
        // Optionally set placeholder text that appears when the mapping is not resolved.
        titleTag.Title = "Book Title";

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(titleTag);

        // Create another content control for the author of the first book.
        StructuredDocumentTag authorTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        authorTag.XmlMapping.SetMapping(xmlPart, "/catalog[1]/book[1]/author[1]", string.Empty);
        authorTag.Title = "Book Author";
        doc.FirstSection.Body.AppendChild(authorTag);

        // Save the resulting document.
        doc.Save("ContentControlXmlBinding.docx");
    }
}
