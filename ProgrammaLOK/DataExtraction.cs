using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ProgrammaLOK
{
    public class DataExtraction
    {
        List<object[]> employee = new List<object[]>();
        List<Vaccination> vaccinations = new List<Vaccination>();
        List<object[]> employeeVaccination = new List<object[]>();

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

        public DataExtraction(string FileName)
        {
            var Rows = getTable(FileName);
        }

        //public object[] e_Table(List<object[]> Rows)
        //{
        //    object[] fio = new object[Rows.Count];
        //    object[] year = new object[Rows.Count];
        //    int i = 0;
        //    foreach (object[] e_row in Rows)
        //    {
        //        do
        //        {
        //            fio[i] = e_row[1];
        //            year[i] = e_row[2];
        //            i++;
        //        }
        //        while (i < 1);
        //    }
        //    return fio;
        //}
    }
}
