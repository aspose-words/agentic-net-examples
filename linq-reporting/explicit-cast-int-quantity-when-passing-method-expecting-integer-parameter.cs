using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Simulate retrieving a decimal value (e.g., from a metered API).
        decimal quantity = 123.45m;

        // Explicitly cast to int when passing to a method that expects an integer.
        SetCustomProperty(doc, "ConsumptionQty", (int)quantity);

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Example method that requires an integer argument.
    static void SetCustomProperty(Document doc, string propertyName, int value)
    {
        var customProps = doc.CustomDocumentProperties;

        if (customProps[propertyName] != null)
            customProps[propertyName].Value = value;
        else
            customProps.Add(propertyName, value);
    }
}
