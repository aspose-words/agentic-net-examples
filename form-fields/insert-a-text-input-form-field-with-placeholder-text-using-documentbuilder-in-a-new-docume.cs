using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field with placeholder text.
            // Parameters: name, type (regular text), format, placeholder text, max length.
            builder.InsertTextInput("CustomerName", TextFormFieldType.Regular, "", "Enter your name", 50);

            // Save the document.
            doc.Save("FormFieldExample.docx");
        }
    }
}
