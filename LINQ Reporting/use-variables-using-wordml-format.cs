using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Access the document builder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some custom variables to the document.
        doc.Variables.Add("CustomerName", "John Doe");
        doc.Variables.Add("OrderNumber", "12345");

        // Insert a DOCVARIABLE field that displays the CustomerName variable.
        FieldDocVariable customerField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        customerField.VariableName = "CustomerName";
        customerField.Update();

        // Insert a line break.
        builder.Writeln();

        // Insert a DOCVARIABLE field that displays the OrderNumber variable.
        FieldDocVariable orderField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        orderField.VariableName = "OrderNumber";
        orderField.Update();

        // Save the document in WordML (XML) format with pretty formatting.
        WordML2003SaveOptions saveOptions = new WordML2003SaveOptions
        {
            SaveFormat = SaveFormat.WordML,
            PrettyFormat = true
        };

        doc.Save("VariablesInWordML.xml", saveOptions);
    }
}
