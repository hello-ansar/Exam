using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace Exam
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            
            InitializeComponent();

            StartPosition = FormStartPosition.CenterScreen;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var items = new Dictionary<string, string>
            {
                { "<KAF>", textBox3.Text },
                { "<KOD>", textBox1.Text },
                { "<DIS>", textBox2.Text },
                { "<DAT>", textBox5.Text },
                { "<NAM>", textBox4.Text },
                { "<NUM>", null },
                { "<Q1>", null },
                { "<Q2>", null },
                { "<Q3>", null }
            };

            var copy = new CopyDoc("C:/Users/Ансар/source/repos/Exam/Exam/шаблон.docx");

            copy.Process(textBox6.Text, items);

            

            //var helper = new Helper("C:/Users/Ансар/source/repos/Exam/Exam/шаблон.docx");



            //helper.Process(items);



        }
    }
}

