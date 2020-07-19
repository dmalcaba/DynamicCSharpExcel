using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DynamicCSharpConsole
{
    public class ReflectionsReview
    {
        public void Start()
        {
            var persons = new List<Person>
            {
                new Person { Id = 1, Name = "George" },
                new Person { Id = 2, Name = "Ringo" },
                new Person { Id = 3, Name = "Paul" },
                new Person { Id = 4, Name = "John" },
            };

            var me = new Person { Id = 5, Name = "Darryl" };
            var isGen = me.GetType().IsGenericType;

            var personType = persons.GetType();

            //Inspect(personType, persons);
            ToObject(persons);


        }

        public void ToObject(object data)
        {
            var type = data.GetType();
            Inspect(type, data);
        }


        public void Inspect(Type type, object data)
        {
            if (type.IsGenericType)
            {
                // determine if it is a list
                string typeName = type.GetGenericTypeDefinition().Name;

                if (typeName == "List`1")
                {
                    // If it is, get the number of elements in the list  
                    int n = (int)type.GetProperty("Count").GetValue(data, null);

                    var name = type.GetGenericArguments()[0].Name;

                    // Process each element in the list  
                    for (int i = 0; i < n; i++)
                    {
                        // Get the list element as type object  
                        object[] index = { i };

                        object myObject = type.GetProperty("Item").GetValue(data, index);
                        Console.WriteLine(myObject.GetType().Name);

                        // Get the object properties  
                        PropertyInfo[] objectProperties = myObject.GetType().GetProperties();

                        // Process each property  
                        foreach (PropertyInfo currentProperty in objectProperties)
                        {
                            string propertyValue = currentProperty.GetValue(myObject, null).ToString();
                            Console.WriteLine("{0} {1}", currentProperty.Name, propertyValue);
                        }

                        // Skip a line between objects  
                        Console.WriteLine();
                    }
                }
            }
        }
    }




    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
