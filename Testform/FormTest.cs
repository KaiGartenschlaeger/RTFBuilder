using System;
using System.Drawing;
using System.Windows.Forms;
using SimpleRtf;

namespace Testform
{
    public partial class FormTest : Form
    {
        bool enableEvent = false;

        public FormTest()
        {
            InitializeComponent();

            RtfBuilder rtf = new RtfBuilder();

            Font font = new Font("Arial", 26);

            for (int i = 0; i < 100; i++)
            {
                Add(rtf, font, Color.Blue, Color.Black);
            }

            richTextBox1.Rtf = rtf.GetRtf();
            richTextBox2.Text = rtf.GetRtf();
            enableEvent = true;
        }

        private void Add(RtfBuilder rtf, Font f, Color c1, Color c2)
        {
            rtf.Append("Kai schreibt", RtfFormat.Bold, c1, f);
            rtf.AppendLine(":", RtfFormat.Bold, c1, f);
            rtf.AppendLine("Dat ist meine Nachricht..", RtfFormat.Normal, c2, f);

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            if (!enableEvent)
            {
                return;
            }
            try
            {
                richTextBox1.Rtf = richTextBox2.Text;
                richTextBox1.BackColor = Color.White;
            }
            catch (Exception)
            {
                richTextBox1.BackColor = Color.Red;
            }
        }
    }
}