using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace ProgrammaLOK
{
    public class Father
    {
        

        public object getTable(object FileName)
        {
            //object FileName = @"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc";
            object rOnly = true;

            //Создаем объект Word - равносильно запуску Word
            Word.Application word = new Word.Application();
            //Создаем документ
            Word.Document doc = null;
            //Открываем документ
            doc = word.Documents.Open(ref FileName, ref rOnly);

            List<object[]> Rows = new List<object[]>();

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
            return Rows;
        }
        public object[] e_Table(List<object[]> Rows)
        {
            object[] fio = new object[Rows.Count];
            object[] year = new object[Rows.Count];
            int i = 0;
            foreach (object[] e_row in Rows)
            {
                do
                {
                    fio[i] = e_row[1];
                    year[i] = e_row[2];
                    i++;
                }
                while (i < 1);
            }
            return fio;
        }
    }
}
