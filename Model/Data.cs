using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ViewAppDocs.Model
{
    internal class Person
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }

        public static Person[] GetPersons()
        {
            //return new Person[5];
            var result = new Person[]
            {
                new Person() { FirstName = "Zanxi", LastName="Zihao"},
                new Person() { FirstName = "linsi", LastName="Mao"},
                new Person() { FirstName = "Rob", LastName="Wirasoro"},
                //new Person() { FirstName = "Ivanov4", LastName="Ivan4"},                
            };

            return result;

        }
    }
}
