using System;
using System.Collections.Generic;
using System.Text;

namespace ProgrammaLOK
{
    public class Employee
    {
        public int idEmployee;
        int yearBirth;
        string phone;

        public Employee(int idEmployee, object yearBirth, object phone)
        {

            

            this.idEmployee = idEmployee;
            int.TryParse(yearBirth == null?"0":yearBirth.ToString(), out this.yearBirth);
            
            this.phone = phone == null?"":phone.ToString();
            //this.phone=phone?.ToString(); //тоже самое что и предыдущая строка только ? => это проверка может ли быть null
            
        }

        public int f(out int h)
        {
            h = 6;
            int t = 7;
            return t;
        }
        //public void Main()
        //{
        //    DataExtraction father = new DataExtraction();
        //    List<object[]> Rows = new List<object[]>();
        //    Rows = (List<object[]>)father.getTable(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
        //    object[] fio = father.e_Table(Rows);
        //}
    }
}
