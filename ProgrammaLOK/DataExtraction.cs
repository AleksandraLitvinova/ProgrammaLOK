using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;

namespace ProgrammaLOK
{
    public class DataExtraction
    {
        List<Employee> employees = new List<Employee>();
        List<Vaccination> vaccinations = new List<Vaccination>();
        List<EmployeeVaccinationRelation> employeesVaccinations = new List<EmployeeVaccinationRelation>();
        List<StatusVaccination> statusesVaccinations = new List<StatusVaccination>();

        public List<object[]> getTable(object FileName)
        {
            //object FileName = @"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc";
            object rOnly = true;

            //Создаем объект Word - равносильно запуску Word
            Word.Application word = new Word.Application();
            //Создаем документ
            Word.Document doc = null;
            //Word.Range range = null;
            //List<object[,]> Rows = new List<object[,]>();
            List<object[]> Rows = new List<object[]>();

            try
            {
                //Открываем документ
                doc = word.Documents.Open(ref FileName, ref rOnly);
                //Word.Table tbl = null;
                //tbl = doc.Tables[1];

                foreach (Word.Table tbl in doc.Tables)
                {
                    for (int i = 3; i <= tbl.Rows.Count; i++)
                    {
                        var row = new object[tbl.Columns.Count];
                        for (int j = 0; j < tbl.Columns.Count; j++)
                        {
                            row[j] = tbl.Cell(i, j + 1).Range.Text;
                        }
                        Rows.Add(row);
                    }
                }
            }
            catch
            {

            }
            finally
            {
                doc.Close();
                doc = null;
                word.Quit();
                word = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            return Rows;
        }

        private void set_employees_list()
        {
            var row_employee = getTable(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");

            int id_count_employee = 1;
            foreach (object[] row in row_employee)
            {
                if (row[1] == null || row[1].ToString() == "")
                    continue;
                Employee emp = new Employee(id_count_employee, row[2], row[3]);
                employees.Add(emp);
                id_count_employee++;
                for (int i = 4; i < row.Length; i++)
                {
                    if (row[i] == null)
                        continue;

                    var cell = row[i].ToString().Trim().Replace("\a", "").Replace("\r", "");
                    if (cell == "")
                        continue;

                    var vac = vaccinations.Find(v => v.id == i - 3);
                    if (vac == null)
                        continue;

                    ////проанализировать статус и взять либо то что в статусе (вакцина) + если есть дата изъять

                    foreach (var val in vaccinations)
                    {
                        int index = cell.ToLower().IndexOf(val.name.ToLower());
                        if (index < 0)
                            continue;

                        vac = val;
                        cell = cell.Substring(val.name.Length + index).Trim();
                        break;
                    }

                    var relation = new EmployeeVaccinationRelation(emp.idEmployee, vac.id);
                    employeesVaccinations.Add(relation);

                    var currentStatus = "Посталена";
                    foreach (var st in statusesVaccinations)
                    {
                        string st_cell = cell;
                        int index = st_cell.ToLower().IndexOf(st.name.ToLower());
                        if (index < 0)
                            continue;

                        //vac = st;
                        st_cell = st_cell.Substring(0, index).Trim()+st_cell.Substring(st.name.Length + index).Trim();
                        if(st_cell !="")
                        {
                            currentStatus = "Отказ";
                        }
                        else
                        {
                            relation.idStatus = get_statusVaccination(cell);
                        }
                        break;
                    }

                    if (int.TryParse(cell, out int year))
                    {
                        relation.dateVaccination = year;
                        relation.idStatus = get_statusVaccination(currentStatus);
                    }
                    else if (DateTime.TryParse(cell, out DateTime data))
                    {
                        relation.dt = data;
                        relation.idStatus = get_statusVaccination(currentStatus);
                    }
                    else if(cell == "")
                    {
                        relation.idStatus = get_statusVaccination(currentStatus);
                    }
                    else
                    {
                            relation.idStatus = get_statusVaccination(cell);
                    }
                }
            }

        }

        private void set_vaccinations_list()
        {
            //var p = new Vaccination(1, "АДС-м RV");
            vaccinations.Add(new Vaccination(1, "АДС-м RV"));
            vaccinations.Add(new Vaccination(2, "Гепатит V1"));
            vaccinations.Add(new Vaccination(3, "Гепатит V2"));
            vaccinations.Add(new Vaccination(4, "Гепатит RV"));
            vaccinations.Add(new Vaccination(5, "Клещ.энцефалит V1"));
            vaccinations.Add(new Vaccination(6, "Клещ.энцефалит V2"));
            vaccinations.Add(new Vaccination(7, "Клещ.энцефалит RV"));
            vaccinations.Add(new Vaccination(8, "Корь"));
            vaccinations.Add(new Vaccination(9, "Краснуха"));
            vaccinations.Add(new Vaccination(10, "АС"));
            vaccinations.Add(new Vaccination(11, "Превенар"));
            vaccinations.Add(new Vaccination(12, "Пневмо-23"));
            vaccinations.Add(new Vaccination(13, "Ковид"));
            vaccinations.Sort((foo1, foo2) => foo2.name.Length.CompareTo(foo1.name.Length));
        }
                    
        private void set_statuses_list()
        {
            statusesVaccinations.Add(new StatusVaccination(1, "Отказ"));
            statusesVaccinations.Add(new StatusVaccination(2, "Поставлено", new List<string>() { "Поставлена", "+" }));
            statusesVaccinations.Add(new StatusVaccination(3, "Болезнь", new List<string>() { "Болела", "Болел" }));
            statusesVaccinations.Add(new StatusVaccination(4, "Антитела", new List<string>() { "а/т +" }));
            statusesVaccinations.Add(new StatusVaccination(5, "Привит", new List<string>() { "Привит", "Привита" }));
            statusesVaccinations.Add(new StatusVaccination(6, "Декретный отпуск", new List<string>() { "Д/о" }));
            statusesVaccinations.Add(new StatusVaccination(7, "Медицинский отвод", new List<string>() { "М/от" }));

        }

        public int get_statusVaccination(string value)
        {
            var st = statusesVaccinations.Find(v => v.Equals(value));

            if (st == null)
            {
                st = new StatusVaccination(statusesVaccinations.Count + 1, value);
                statusesVaccinations.Add(st);
            }

            return st.id;
        }

        public DataExtraction(string FileName)
        {
            set_vaccinations_list();
            set_statuses_list();
            set_employees_list();

            string statuses = "";
            foreach (StatusVaccination sVac in statusesVaccinations)
            {
                if (statuses != "")
                    statuses += Environment.NewLine;
                statuses += sVac.name;
            }
        }

    }
}
