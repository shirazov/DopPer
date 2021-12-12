using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using MaterialSkin.Controls;
using MaterialSkin;
using Microsoft.Office.Interop.Word;
using MySql.Data;
using MySql.Data.MySqlClient;
using DataTable = System.Data.DataTable;

namespace DopPer
{
    public partial class DopForm : MaterialForm
    {
        #region Doc
        public Document doc { get; private set; }
        public Word.Application wordapp1 { get; private set; }
        #endregion

        public DopForm()
        {
            InitializeComponent();

            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Green800, Primary.Green900, Primary.Green500, Accent.Green400, TextShade.WHITE);

        }

        private void Zap_v_stats()
        {

            using (MySqlConnection mycon = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root;charset=utf8"))
            {
                mycon.Open();
                string sql_registrUser = "INSERT INTO `doper`.`stats` ( `forma`, `spes`, `gruppa`, `kursSemestr`, `disip`, `fio_stud`, `fio_prepod`, `time`) VALUES( '" + formCont.Text + "', '" + spes.Text + "', '" + clas.Text + "', '" + curs.Text + "/" + semestr.Text + "', '" + lessen.Text + "', '" + student.Text + "', '" + prepod.Text + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "');";
                MySqlCommand comm_registr = new MySqlCommand(sql_registrUser, mycon);
                comm_registr.ExecuteNonQuery();
                mycon.Close();
            }
        }

        private void Peredacha()
        {
            ProgressBarGlav.Value = ProgressBarGlav.Value + 5;

            try
            {
                // Создаём объект приложения
                Word.Application app = new Word.Application();

                // Путь до шаблона документа
                string source = @"C:\Users\shirz\Desktop\DopPer.docx"; //

                // Открываем
                doc = app.Documents.Open(source);
                doc.Activate();
                ProgressBarGlav.Value = ProgressBarGlav.Value + 15;
                // Добавляем информацию
                // wBookmarks содержит все закладки
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;
                string[] data = new string[8] { clas.Text, curs.Text, formCont.Text, lessen.Text, prepod.Text, semestr.Text, spes.Text, "888" };
                foreach (Word.Bookmark mark in wBookmarks)
                {

                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }
                ProgressBarGlav.Value = ProgressBarGlav.Value + 30;
                // Закрываем документ
                doc.Close();
                doc = null;
                ProgressBarGlav.Value = ProgressBarGlav.Value + 15;
            }
            catch (Exception ex) // c выводом ошибок нужно разобраться, исключения всё крашат
            {
                try
                {
                    // Если произошла ошибка, то
                    // закрываем документ и выводим информацию
                    MessageBox.Show(ex.Message);
                    doc.Close();
                    doc = null;
                    //Console.WriteLine("Во время выполнения произошла ошибка!");
                    //Console.ReadLine();
                }
                catch
                {
                    MessageBox.Show(ex.ToString(), "Эх", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            ProgressBarGlav.Value = ProgressBarGlav.Value + 15;
            Zap_v_stats();
            ProgressBarGlav.Value = ProgressBarGlav.Value + 20;
        }

        private void setButton_Click(object sender, EventArgs e)
        {
            Peredacha();
        }
        #region funs
        private void Stats()
        {

            MySqlConnection mycon = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root");
            mycon.Open();

            string sql = "SELECT * FROM doper.stats;";
            // Создать объект Command.
            MySqlCommand cmd = new MySqlCommand();
            cmd.Connection = mycon;

            cmd.CommandText = sql;
            MySqlDataReader reader = cmd.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            dataGridVie.DataSource = table;
        }
        private void FillComboBox()
        {
            try
            {


                MySqlConnection mycon = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root");
                mycon.Open();


                string s = "SELECT formcontcol AS name_form, idform AS id_form FROM doper.formcont;";
                MySqlCommand mcd = new MySqlCommand(s, mycon);
                DataTable formTable = new DataTable();
                MySqlDataAdapter adapter = new MySqlDataAdapter(mcd);
                try
                {
                    adapter.Fill(formTable);
                    formCont.DataSource = formTable;
                    formCont.DisplayMember = "name_form";
                    formCont.ValueMember = "id_form";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "22222222", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                string s1 = "SELECT idspes AS id_spes, spes_name AS name_spes FROM doper.spes;";
                MySqlCommand mcd1 = new MySqlCommand(s1, mycon);
                DataTable spesTable = new DataTable();
                MySqlDataAdapter adapter1 = new MySqlDataAdapter(mcd1);
                try
                {
                    adapter1.Fill(spesTable);
                    spes.DataSource = spesTable;
                    spes.DisplayMember = "name_spes";
                    spes.ValueMember = "id_spes";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "111111111", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                string s4 = "SELECT * FROM doper.lecturers;";
                MySqlCommand mcd4 = new MySqlCommand(s4, mycon);
                MySqlDataReader mdr4 = mcd4.ExecuteReader();
                while (mdr4.Read())
                {
                    prepod.Items.Add(mdr4.GetString("name"));
                }
                mdr4.Close();

                mycon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "MySQL Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region tab4
        private MySqlConnection mySqlConnection1 = null;
        private MySqlCommandBuilder mySqlCommandBuilder1 = null;
        private MySqlDataAdapter mySqlDataAdapter1 = null;
        private DataSet dataSet1 = null;
        private bool newRowAdding = false;

        private void Loadtab4()
        {
            try
            {
                mySqlDataAdapter1 = new MySqlDataAdapter("SELECT *, 'Delete' AS Действие FROM class", mySqlConnection1);
                mySqlCommandBuilder1 = new MySqlCommandBuilder(mySqlDataAdapter1);

                dataSet1 = new DataSet();

                mySqlCommandBuilder1.GetInsertCommand();
                mySqlCommandBuilder1.GetUpdateCommand();
                mySqlCommandBuilder1.GetDeleteCommand();

                mySqlDataAdapter1.Fill(dataSet1, "class");

                dataGridView1.DataSource = dataSet1.Tables["class"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, i] = linkCell;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Reloadtab4()
        {
            try
            {
                dataSet1.Tables["class"].Clear();

                mySqlDataAdapter1.Fill(dataSet1, "class");

                dataGridView1.DataSource = dataSet1.Tables["class"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, i] = linkCell;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Reloadtab4();
        }

        private void dataGridV3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 4)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridView1.Rows.RemoveAt(rowIndex);

                            dataSet1.Tables["class"].Rows[rowIndex].Delete();

                            mySqlDataAdapter1.Update(dataSet1, "class");
                        }
                    }
                    else if (task == "Insert")
                    {
                        int rowIndex = dataGridView1.Rows.Count - 2;
                        DataRow row = dataSet1.Tables["class"].NewRow();

                        row["idclass"] = dataGridView1.Rows[rowIndex].Cells["idclass"].Value;
                        row["class_name"] = dataGridView1.Rows[rowIndex].Cells["class_name"].Value;
                        row["idspes"] = dataGridView1.Rows[rowIndex].Cells["idspes"].Value;
                        row["idgroups"] = dataGridView1.Rows[rowIndex].Cells["idgroups"].Value;

                        dataSet1.Tables["class"].Rows.Add(row);

                        dataSet1.Tables["class"].Rows.RemoveAt(dataSet1.Tables["class"].Rows.Count - 1);

                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);

                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "Удалить";

                        mySqlDataAdapter1.Update(dataSet1, "class");

                        newRowAdding = false;
                    }
                    else if (task == "Update")
                    {
                        int r = e.RowIndex;

                        dataSet1.Tables["class"].Rows[r]["idclass"] = dataGridView1.Rows[r].Cells["idclass"].Value;
                        dataSet1.Tables["class"].Rows[r]["class_name"] = dataGridView1.Rows[r].Cells["class_name"].Value;
                        dataSet1.Tables["class"].Rows[r]["idspes"] = dataGridView1.Rows[r].Cells["idspes"].Value;
                        dataSet1.Tables["class"].Rows[r]["idgroups"] = dataGridView1.Rows[r].Cells["idgroups"].Value;

                        mySqlDataAdapter1.Update(dataSet1, "class");

                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "Delete";
                    }

                    Reloadtab4();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridV3_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    newRowAdding = true;

                    int lastRow = dataGridView1.Rows.Count - 2;

                    DataGridViewRow row = dataGridView1.Rows[lastRow];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, lastRow] = linkCell;

                    row.Cells["Действие"].Value = "Insert";


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridV3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;

                    DataGridViewRow editinRow = dataGridView1.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[4, rowIndex] = linkCell;

                    editinRow.Cells["Действие"].Value = "Update";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void spes_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            MySqlConnection mycon = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root");
            mycon.Open();
            clas.Text = "";
            string s2 = "SELECT idgroups AS id_clas, class_name AS name_clas FROM doper.class WHERE idspes=" + spes.SelectedValue + ";";

            MySqlCommand mcd2 = new MySqlCommand(s2, mycon);
            DataTable clasTable = new DataTable();
            MySqlDataAdapter adapter2 = new MySqlDataAdapter(mcd2);
            try
            {
                adapter2.Fill(clasTable);
                clas.DataSource = clasTable;
                clas.DisplayMember = "name_clas";
                clas.ValueMember = "id_clas";
            }
            catch
            {

            }
        }

        private void clas_SelectedIndexChanged(object sender, EventArgs e)
        {
            MySqlConnection mycon = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root");
            mycon.Open();
            lessen.Text = "";
            string s3 = "SELECT ID AS id_dis, name AS name_dis FROM doper.dis WHERE group_id=" + clas.SelectedValue + ";";

            MySqlCommand mcd3 = new MySqlCommand(s3, mycon);
            DataTable lessebTable = new DataTable();
            MySqlDataAdapter adapter3 = new MySqlDataAdapter(mcd3);
            try
            {
                adapter3.Fill(lessebTable);
                lessen.DataSource = lessebTable;
                lessen.DisplayMember = "name_dis";
                lessen.ValueMember = "id_dis";
            }
            catch
            {

            }

            student.Text = "";
            string s4 = "SELECT ID AS id_stud, FIO AS name_stud FROM doper.students WHERE group_id=" + clas.SelectedValue + ";";

            MySqlCommand mcd4 = new MySqlCommand(s4, mycon);
            DataTable studentbTable = new DataTable();
            MySqlDataAdapter adapter4 = new MySqlDataAdapter(mcd4);
            try
            {
                adapter4.Fill(studentbTable);
                student.DataSource = studentbTable;
                student.DisplayMember = "name_stud";
                student.ValueMember = "id_stud";
            }
            catch
            {

            }
        }

        private void dataGridV3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);

            //вместо цифры ячейка где только числовые
            if (dataGridView1.CurrentCell.ColumnIndex == 0)
            {
                TextBox textBox = e.Control as TextBox;

                if (textBox != null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
            if (dataGridView1.CurrentCell.ColumnIndex == 2)
            {
                TextBox textBox = e.Control as TextBox;

                if (textBox != null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
            if (dataGridView1.CurrentCell.ColumnIndex == 3)
            {
                TextBox textBox = e.Control as TextBox;

                if (textBox != null)
                {
                    textBox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
        }
        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        #endregion


        private void DopForm_Load(object sender, EventArgs e)
        {
            FillComboBox();
            Stats();

            mySqlConnection1 = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root;charset=utf8");
            mySqlConnection1.Open();
            Loadtab4();
        }
  
        private void materialButton1_Click(object sender, EventArgs e)
        {

            MySqlConnection mycon = new MySqlConnection("Server=localhost;port=3306;Database=doper;Uid=root;Pwd=root;username=root;password=root");
            mycon.Open();

            dt1.Text = datetime_1.Text; dt2.Text = datetime_2.Text;

            string sql = "SELECT * FROM doper.stats WHERE time BETWEEN '"+ datetime_1.Text + "' AND '" + datetime_2.Text + "';";
            // Создать объект Command.
            MySqlCommand cmd = new MySqlCommand();
            cmd.Connection = mycon;

            cmd.CommandText = sql;
            MySqlDataReader reader1 = cmd.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader1);
            dataGridVie.DataSource = table;
        }

        private void materialButton2_Click(object sender, EventArgs e)
        {
            Stats();
        }

        private void Sbros()
        {
            ProgressBarGlav.Value = 0;
            formCont.Text = "";
            spes.Text = "";
            clas.Text = "";
            lessen.Text = "";
            student.Text = "";
            prepod.Text = "";
        }

        private void materialButton2_Click_1(object sender, EventArgs e)
        {
            Sbros();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
    }
}
