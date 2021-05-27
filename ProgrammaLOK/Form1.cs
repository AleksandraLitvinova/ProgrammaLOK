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

            Employee emp = new Employee();
            emp.getTable();
            emp.e_Table();
        }

        
        

        //Function f = new Function();
        //f.doAfter = Program.Method;
        //f.doSomething();

    }
}
