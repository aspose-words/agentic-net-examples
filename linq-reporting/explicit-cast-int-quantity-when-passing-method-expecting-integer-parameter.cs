using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path).
        Document doc = new Document("input.docx");

        // Retrieve a decimal value from the metered licensing API.
        decimal quantity = Aspose.Words.Metered.GetConsumptionQuantity();

        // The method SetCustomProperty expects an int, so we cast explicitly.
        SetCustomProperty(doc, "ConsumptionQty", (int)quantity);

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Example method that requires an integer argument.
    static void SetCustomProperty(Document doc, string propertyName, int value)
    {
        // Access the custom document properties collection.
        var customProps = doc.CustomDocumentProperties;

        // If the property already exists, update its value; otherwise, add a new one.
        if (customProps[propertyName] != null)
            customProps[propertyName].Value = value;
        else
            customProps.Add(propertyName, value);
    }
}
