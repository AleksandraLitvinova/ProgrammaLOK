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

        public List<object[,]> getTable(object FileName)
        {
            //object FileName = @"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc";
            object rOnly = true;

            //Создаем объект Word - равносильно запуску Word
            Word.Application word = new Word.Application();
            //Создаем документ
            Word.Document doc = null;
            //Word.Range range = null;
            List<object[,]> Rows = new List<object[,]>();

            try
            {
                //Открываем документ
                doc = word.Documents.Open(ref FileName, ref rOnly);
                Word.Table tbl = null;



                //for (int i = 1; i<doc.Characters.Count; i++)
                //{
                    tbl = doc.Tables[1];
                    object begCell = tbl.Cell(3, 1).Range.Start;
                    object endCell = tbl.Cell(tbl.Rows.Count, 13).Range.End;
                    Word.Range wordcellrange = doc.Range(ref begCell, ref endCell);

                var t =wordcellrange.value

                    object[,] mas = new object[tbl.Rows.Count-2, 12];
                for(int i=0;i<tbl.Rows.Count;i++)
                {
                    for (int j = 0; j < 12; j++)
                    {
                        wordcellrange = tbl.Cell(i+3, j + 1).Range;
                        mas[i, j] = wordcellrange.Text;
                    }
                    
                    Rows.Add(mas);
                }
                //}



                //foreach (Word.Table tbl in doc.Tables)
                //{               
                //for (int i = 3; i <= tbl.Rows.Count; i++)
                //    {
                //        var row = new object[tbl.Columns.Count];
                //        for (int j = 0; j < tbl.Columns.Count; j++)
                //        {
                //            row[j] = tbl.Cell(i, j + 1).Range.Text;
                //        }
                //        Rows.Add(row);
                //    }
                ////}
            }
            catch {

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

        //private void set_employees_list()
        //{
        //    var row_employee = getTable(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
        //    int id_count_employee = 1;
        //    foreach (object[] row in row_employee)
        //    {
        //        if (row[1] == null || row[1].ToString() == "")
        //            continue;
        //        Employee emp = new Employee(id_count_employee, row[2], row[3]);
        //        employees.Add(emp);
        //        id_count_employee++;
        //        for (int i = 4; i < row.Length; i++)
        //        {
        //            if (row[i] == null)
        //                continue;
        //            var cell = row[i].ToString().Trim().Replace("\a", "").Replace("\r", "");
        //            if (cell == "")
        //                continue;

        //            var vac = vaccinations.Find(v => v.id == i - 3);
        //            if (vac == null)
        //                continue;
        //            var relation = new EmployeeVaccinationRelation(emp.idEmployee, vac.id);
        //            employeesVaccinations.Add(relation);
        //            if (int.TryParse(cell, out int data))
        //            {
        //                relation.dateVaccination = data;
        //                relation.idStatus = get_statusVaccination("Посталена");
                        
        //            }
        //            else if (DateTime.TryParse(cell, out DateTime data2))
        //            {
        //                relation.dt = data2;
        //                relation.idStatus = get_statusVaccination("Посталена");
        //            }
        //            else
        //            {
        //                relation.idStatus = get_statusVaccination(cell);
        //            }

        //        }
        //    }
        //}

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
        }

        public void set_employeesVaccinations_list()
        {


        }

        public int get_statusVaccination(string value)
        {
            var st = statusesVaccinations.Find(v => v.name == value);

            if (st == null)
            {
                st = new StatusVaccination(statusesVaccinations.Count, value);
                statusesVaccinations.Add(st);
            }

            return st.id;
        }

        public DataExtraction(string FileName)
        {
            var Rows = getTable(FileName);
            set_vaccinations_list();
            //set_employees_list();
        }

    }
}
