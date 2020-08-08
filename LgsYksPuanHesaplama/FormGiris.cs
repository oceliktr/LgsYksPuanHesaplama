using System;
using System.Windows.Forms;

namespace LgsYksPuanHesaplama
{
    public partial class FormGiris : Form
    {
        public FormGiris()
        {
            InitializeComponent();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            FormLgs frm = new FormLgs();
            frm.Show();
            this.Hide();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            FormYks frm = new FormYks();
            frm.Show();
            this.Hide();
        }

        private void FormGiris_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
