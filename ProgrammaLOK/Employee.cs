using System;
using System.Collections.Generic;
using System.Text;

namespace ProgrammaLOK
{
    public class Employee
    {
        int idEmployee;
        int yearBirth;
        int phone;

        public void Main()
        {
            Father father = new Father();
            List<object[]> Rows = new List<object[]>();
            Rows = (List<object[]>)father.getTable(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
            object[] fio = father.e_Table(Rows);
        }
    }
}
