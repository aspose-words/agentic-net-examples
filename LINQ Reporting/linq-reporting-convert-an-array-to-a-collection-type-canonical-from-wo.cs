using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document.
            Document doc = new Document();

            // 2. Build a template that uses LINQ Reporting syntax.
            //    The template will iterate over a collection named "persons"
            //    and output each person's Name and Age.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("<<foreach [persons]>>");
            builder.Writeln("Name: <<[Name]>>, Age: <<[Age]>>");
            builder.Writeln("<</foreach>>");

            // 3. Prepare an array of Person objects.
            Person[] personArray = new Person[]
            {
                new Person("John Doe", 30),
                new Person("Jane Smith", 25),
                new Person("Bob Johnson", 40)
            };

            // 4. The ReportingEngine expects a collection that implements IEnumerable.
            //    An array already implements IEnumerable, but to demonstrate conversion
            //    to a canonical collection type we explicitly call ToArray on the array.
            //    (In real scenarios you might convert a List<T> to an array, etc.)
            Person[] canonicalArray = personArray.ToArray();

            // 5. Populate the template with the data source.
            //    The third argument is the name used in the template to reference the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, canonicalArray, "persons");

            // 6. Save the resulting document.
            doc.Save("LinqReportingArrayToCollection.docx");
        }
    }
}
