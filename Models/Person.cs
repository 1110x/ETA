using System.Collections.ObjectModel;  // ← 이 줄 추가!

namespace ETA.Models  // 또는 당신 namespace
{
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int Age { get; set; }
        public ObservableCollection<Person> Children { get; } = new();  // 이제 OK

        public Person(string first, string last, int age)
        {
            FirstName = first;
            LastName = last;
            Age = age;
        }
    }
}