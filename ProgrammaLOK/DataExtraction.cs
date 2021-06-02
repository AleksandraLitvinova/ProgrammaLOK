using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

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
            List<object[]> Rows = new List<object[]>();

            try
            {
                //Открываем документ
                doc = word.Documents.Open(ref FileName, ref rOnly);
                Word.Table tbl = doc.Tables[1];
                
                //foreach (Word.Table tbl in doc.Tables)
                //{               
                for (int i = 3; i <= tbl.Rows.Count; i++)
                {
                    var row = new object[tbl.Columns.Count];
                    for (int j = 0; j < tbl.Columns.Count; j++)
                    {
                        row[j] = tbl.Cell(i, j + 1).Range.Text;
                    }

                    Rows.Add(row);
                }
                //}
            }
            catch{

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

        public void set_employees_list()
        {
            var row_employee = getTable(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
            int id_count_employee = 1;
            foreach(object[] row in row_employee)
            {
                if(row[1] == null || row[1].ToString() == "")
                    continue;
                Employee emp = new Employee(id_count_employee, row[2], row[3]);
                employees.Add(emp);
                id_count_employee++;
                for(int i=4; i<row.Length;i++)
                {
                    if (row[i] == null)
                        continue;
                    var cell = row[i].ToString().Trim();
                    if (cell == "")
                        continue;

                    var vac = vaccinations.Find(v=>v.id==i-3);
                    if (vac == null)
                        continue;
                    var relation = new EmployeeVaccinationRelation(emp.idEmployee, vac.id);
                    employeesVaccinations.Add(relation);
                    if(int.TryParse(cell, out int data))
                    {
                        relation.dateVaccination = data;
                    }
                        
                }
            }
        }

        public void set_vaccinations_list()
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
        }

        public void set_employeesVaccinations_list()
        {
            

        }

        public object set_statusesVaccinations_list(object[] row)
        {
            var rows_employee = getTable(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
            int id_count_satus = 1;
            //string name_st;
            //StatusVaccination sv = new StatusVaccination(id_count_satus, );
            for(int i=4; i<row.Length;i++)
            {
                var cell = row[i].ToString().Trim();
                if (int.TryParse(cell, out int data))
                    continue;
                else
                {
                    statusesVaccinations.Add(new StatusVaccination(id_count_satus, cell));
                }
            }
            



            statusesVaccinations.Add(new StatusVaccination(1, "Отказ"));
            statusesVaccinations.Add(new StatusVaccination(2, "Переболел/а"));
            statusesVaccinations.Add(new StatusVaccination(3, "Антитела"));
            statusesVaccinations.Add(new StatusVaccination(4, "Привит/а"));
            statusesVaccinations.Add(new StatusVaccination(5, "Декретный отпуск"));
            statusesVaccinations.Add(new StatusVaccination(6, "Мед отвод"));
            statusesVaccinations.Add(new StatusVaccination(7, "+"));
        }

        public DataExtraction(string FileName)
        {
            var Rows = getTable(FileName);
        }

    }
}
