using System;
using System.Collections.Generic;
using System.Text;

namespace ProgrammaLOK
{
    class Vaccination
    {
        public int id;
        public string name;
        int period;
        int age;
        int idNext;

        public Vaccination(int id, string name, int period, int age, int idNext)
        {
            this.id = id;
            this.name = name;
            this.period = period;
            this.age = age;
            this.idNext = idNext;
        }

        public Vaccination(int id, string name)
        {
            this.id = id;
            this.name = name;
        }
    }
}