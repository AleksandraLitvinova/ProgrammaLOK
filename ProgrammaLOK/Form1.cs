using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ProgrammaLOK
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            int index = 0;
            //for (int j = 0; j < st.Length; j++)
            //{
            //    index = cell.ToLower().IndexOf(st2);
            //    if (index >= 0)
            //    {
            //        switch (j)
            //        {
            //            case 0:
            //                id_pn = vac.id + 7;
            //                break;
            //            case 1:
            //                id_pn = vac.id + 8;
            //                break;
            //            case 2:
            //                id_pn = vac.id + 9;
            //                break;
            //            case 3:
            //                id_pn = vac.id + 10;
            //                break;
            //        }
            //        cell = cell.Substring(st[j].Length + index).Trim();
            //        //int id_pn = vac.id + 7;
            //        relation = new EmployeeVaccinationRelation(emp.idEmployee, id_pn);
            //        employeesVaccinations.Add(relation);
            //        break;
            //    }
            //}
            DataExtraction d = new DataExtraction(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}