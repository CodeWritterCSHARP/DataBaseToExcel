using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace DataBase
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int counter = -1;
        Dictionary<string, Tuple<string, string, DateTime>> dict = new Dictionary<string, Tuple<string, string, DateTime>>();

        private void button1_Click(object sender, EventArgs e)
        {
            #region tarkistus
            string ekaosa = "";
            bool addingtodict = true;

            try { for (int i = 0; i < 6; i++) ekaosa += textBox1.Text[i]; }
            catch { addingtodict = false; MessageBox.Show("Henkilötunnuksen ekaosa puuttuu"); }

            if(addingtodict == true)
            {
                try { int p = Convert.ToInt32(ekaosa); }
                catch { addingtodict = false; MessageBox.Show("Henkilötunnuksessa on virhe"); }
            }

            if (string.IsNullOrEmpty(textBox2.Text)) { addingtodict = false; MessageBox.Show("Nimi puuttuu"); }
            if (string.IsNullOrEmpty(textBox3.Text)) { addingtodict = false; MessageBox.Show("Puhelinnumero puuttuu"); }

            if (textBox1.Text.Length != 11 && addingtodict == true) { addingtodict = false; MessageBox.Show("Henkilötunnuksen pituus pitää olla 11 merkkiä"); }
            if (textBox3.Text.Length > 17 && addingtodict == true) { addingtodict = false; MessageBox.Show("Puhelinnumeron pituus ei voi olla > 17"); }

            if(addingtodict == true)
            {
                int count = 0;
                for (int i = 0; i < textBox3.Text.Length; i++)
                    if (Char.IsDigit(textBox3.Text[i])) count++;
                if (count > 12) { addingtodict = false; MessageBox.Show("Puhelinnumerossa on liian paljon numeroita, max määrä on 12"); }
                if (count < 10) { addingtodict = false; MessageBox.Show("Puhelinnumerossa on numeroiden pula (järin paljoa), min määrä on 10"); }
            }
            #endregion

            #region addingintable
            if(addingtodict == true)
            {
                bool addingtotable = true;
                DateTime dt = DateTime.Now;

                try { dict.Add(textBox1.Text, Tuple.Create(textBox2.Text, textBox3.Text, dt)); }
                catch { addingtotable = false; MessageBox.Show("Tunnus jo on taulukossa"); }

                if(addingtotable == true)
                {
                    this.dataGridView1.Rows.Add();
                    counter++;
                    label11.Text = (counter + 1).ToString();
                    dataGridView1.Rows[counter].Cells["Henkilötunnus"].Value = (object)dict.ElementAt(counter).Key;
                    dataGridView1.Rows[counter].Cells["Nimi"].Value = (object)dict.ElementAt(counter).Value.Item1;
                    dataGridView1.Rows[counter].Cells["Puhelinnumero"].Value = (object)dict.ElementAt(counter).Value.Item2;
                    dataGridView1.Rows[counter].Cells["Päivitys"].Value = (object)dict.ElementAt(counter).Value.Item3;
                }
            }
            #endregion
        }

        private void button3_Click(object sender, EventArgs e)
        {
            #region savingtoexel
            if(dataGridView1.RowCount > 1)
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel.xlxs|*.xlxs| Excel.xlsm|*.xlsm" })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK) textBox5.Text = openFileDialog.FileName;
                    try
                    {
                        _Application app = new Microsoft.Office.Interop.Excel.Application();
                        _Workbook workbook = app.Workbooks.Add(Type.Missing);
                        _Worksheet worksheet = null;
                        app.Visible = false;
                        worksheet = workbook.Sheets["Taul1"];
                        for (int i = 1; i < dataGridView1.ColumnCount + 1; i++) worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {
                                if (dataGridView1.Rows[i].Cells[j].Value != null)
                                {
                                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                }
                                else
                                {
                                    worksheet.Cells[i + 2, j + 1] = "";
                                }
                            }
                        }
                        try
                        {
                            workbook.SaveAs(textBox5.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        }
                        catch (Exception ex1)
                        {
                            MessageBox.Show(ex1.Message);
                        }
                        finally { app.Quit(); }
                    }
                    catch (Exception ex1) { MessageBox.Show(ex1.Message); }
                }
            }
            else { MessageBox.Show("Taulukpssa ei oo mitään"); }
            #endregion
        }

        private void button2_Click(object sender, EventArgs e)
        {
            #region deletingfromtable
            int current = 0;
            int k = 0;
            bool checker = true;

            if (dict.ContainsKey(textBox4.Text))
            {
                dict.Remove(textBox4.Text);
                counter--;
                label11.Text = (counter + 1).ToString();

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    if(Convert.ToString(dataGridView1.Rows[i].Cells["Henkilötunnus"].Value) == textBox4.Text)
                    {
                        current = i;
                        for (int j = 0; j < dataGridView1.ColumnCount; j++) dataGridView1.Rows[i].Cells[j].Value = "";
                        break;
                    }
                }
            }
            else { MessageBox.Show("Ohjelma ei löytänyt henkilötunnusta"); checker = false; }

            if(checker == true)
            {
                while (string.IsNullOrEmpty(Convert.ToString(dataGridView1.Rows[k].Cells[0].Value)) == true) k++;
                for (int i = current; i < dataGridView1.RowCount - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        dataGridView1.Rows[i].Cells[j].Value = dataGridView1.Rows[i + 1].Cells[j].Value;
                }

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    if(string.IsNullOrEmpty(Convert.ToString(dataGridView1.Rows[i].Cells[0].Value)) == true) { dataGridView1.Rows.Remove(dataGridView1.Rows[i]); }
                #endregion

            #region dictionaryrebuilding
                dict.Clear();
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    dict.Add(dataGridView1.Rows[i].Cells["Henkilötunnus"].Value.ToString(), Tuple.Create(dataGridView1.Rows[i].Cells["Nimi"].Value.ToString(),
                        dataGridView1.Rows[i].Cells["Puhelinnumero"].Value.ToString(), Convert.ToDateTime(dataGridView1.Rows[i].Cells["Päivitys"].Value)));
                }
            }
            #endregion
        }

        private void button4_Click(object sender, EventArgs e)
        {
            #region findingmatches
            int c = 0;
            label8.Text = "Taulukpssa on ";
            label9.Text = "Taulukpssa on ";

            if (!string.IsNullOrEmpty(textBox7.Text))
            {
                c = dict.Values.Count(x => x.Item1 == textBox7.Text);
                label8.Text += c.ToString() + " nimeä";
                c = 0;
            }

            if (!string.IsNullOrEmpty(textBox6.Text))
            {
                c = dict.Values.Count(x => x.Item2 == textBox6.Text);
                label9.Text += c.ToString() + " puhelinnumeroa";
            }
            #endregion
        }
    }
}
