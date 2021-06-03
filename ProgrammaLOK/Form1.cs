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

            DataExtraction d = new DataExtraction(@"\\Devsrv\dtd\Материалы\Материалы для проектов\ПП для ЛОК\Прививки(копия).doc");
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }




        //Function f = new Function();
        //f.doAfter = Program.Method;
        //f.doSomething();

    }
}
