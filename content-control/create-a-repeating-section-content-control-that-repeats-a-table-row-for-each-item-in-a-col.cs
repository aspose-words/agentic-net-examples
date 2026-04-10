using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.EndRow();

        // Prepare sample data as XML and add it as a custom XML part.
        string xml = @"<people>
                         <person><name>John Doe</name><age>30</age></person>
                         <person><name>Jane Smith</name><age>25</age></person>
                         <person><name>Bob Johnson</name><age>40</age></person>
                       </people>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add("People", xml);

        // Create a repeating section content control at the row level.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        // Map the repeating section to the collection of <person> elements.
        repeatingSection.XmlMapping.SetMapping(xmlPart, "/people[1]/person", string.Empty);
        table.AppendChild(repeatingSection);

        // Create a repeating section item (the template row that will be repeated).
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
        repeatingSection.AppendChild(repeatingItem);

        // Build the row that will be cloned for each <person>.
        Row dataRow = new Row(doc);
        repeatingItem.AppendChild(dataRow);

        // Cell for the person's name.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        nameSdt.XmlMapping.SetMapping(xmlPart, "/people[1]/person[1]/name[1]", string.Empty);
        dataRow.AppendChild(nameSdt);

        // Cell for the person's age.
        StructuredDocumentTag ageSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        ageSdt.XmlMapping.SetMapping(xmlPart, "/people[1]/person[1]/age[1]", string.Empty);
        dataRow.AppendChild(ageSdt);

        // End the table.
        builder.EndTable();

        // Save the resulting document.
        doc.Save("RepeatingSectionTable.docx");
    }
}
