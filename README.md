using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;




namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        OleDbConnection con;
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;
        DataTable dt;

        OleDbConnection con1;
        OleDbDataAdapter da1;
        OleDbCommand cmd1;
        DataSet ds1;
        DataTable dt1;


        //여신등록
        OleDbConnection con2;
        OleDbDataAdapter da2;
        OleDbCommand cmd2;
        DataSet ds2;
        DataTable dt2;







        public Form1()
        {
            InitializeComponent();
        }

        //가상계좌등록
        void GetAccount()
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=vslist.accdb");
            da = new OleDbDataAdapter("SELECT *FROM Sheet1", con);
            ds = new DataSet();
            dt = new DataTable();

            con.Open();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }

        //매크로등록
        void macro()
        {
            con1 = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=macro.accdb");
            da1 = new OleDbDataAdapter("SELECT *FROM macro", con1);
            ds1 = new DataSet();
            dt1 = new DataTable();

            con1.Open();
            da1.Fill(dt1);
            dataGridView2.DataSource = dt1;
            con1.Close();
        }

        //여신등록
        void debt()
        {
            con2= new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=debt.accdb");
            da2 = new OleDbDataAdapter("SELECT *FROM debt", con2);
            ds2 = new DataSet();
            dt2 = new DataTable();

            con2.Open();
            da2.Fill(dt2);
            dataGridView4.DataSource = dt2;
            con2.Close();

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            GetAccount();
            macro();
            debt();


        }


        private void 프로그램종료XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }



        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.SelectedRows != null)
            {
                richTextBox2.Text = "입금계좌 안내드리겠습니다. \r\n" + "\r\n"
                    + "1.1회 최대입금금액은 100만원입니다.\r\n"
                    + "2.1회 100만원 이상 입금시 충전처리가 불가함을 말씀드립니다.\r\n"
                    + "3.회원정보에 기재하신 정보외에는 입금처리가 불가합니다.\r\n" + "\r\n"
                    + "경남은행" + "\n" + dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();

                richTextBox6.Text = "입금계좌 안내드리겠습니다. \r\n" + "\r\n"
                    + "1.1회 최대입금금액은 300만원 / 1일 최대 입금금액은 600만원입니다.\r\n"
                    + "2.1일 600만원 이상 입금시 충전처리가 불가합니다.\r\n"
                    + "3.회원정보에 기재하신 정보외에는 입금처리가 불가합니다..\r\n" + "\r\n"
                    + "경남은행" + "\n" + dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();

                richTextBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                richTextBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                comboBox2.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            }
        }

        private void buttonXP2_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox2.Text);
        }


        private void buttonXP1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.richTextBox1.Text))
            {


            }
            else
            {

                DataView dv = new DataView(dt, "월렛사용자명 = '" + this.richTextBox1.Text + "'", "월렛사용자명 asc", DataViewRowState.CurrentRows);
                dataGridView1.DataSource = dv;
            }
            richTextBox1.Text = string.Empty;
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                buttonXP1_Click(sender, e);

            e.Handled = true;
            e.SuppressKeyPress = true;

        }
        //계좌등록
        private void buttonXP3_Click(object sender, EventArgs e)
        {
            string query = "Insert into Sheet1 (월렛사용자명,가상계좌은행,가상계좌번호,이용사,등록일) values (@Name,@Bank,@Vaccount,@Company,@Rdate)";
            cmd = new OleDbCommand(query, con);
            cmd.Parameters.AddWithValue("@Name", richTextBox3.Text);
            cmd.Parameters.AddWithValue("@Bank", richTextBox4.Text);
            cmd.Parameters.AddWithValue("@Vaccount", richTextBox5.Text);
            cmd.Parameters.AddWithValue("@Company", comboBox2.SelectedItem);
            cmd.Parameters.AddWithValue("@Rdate", dateTimePicker2.Value);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            GetAccount();
        }


        //매크로등록
        private void buttonXP8_Click(object sender, EventArgs e)
        {
            string macroquery = "Insert into macro (구분,내용) values (@Division,@Context)";
            cmd1 = new OleDbCommand(macroquery, con1);
            cmd1.Parameters.AddWithValue("@Division", richTextBox9.Text);
            cmd1.Parameters.AddWithValue("@Context", richTextBox7.Text);


            con1.Open();
            cmd1.ExecuteNonQuery();
            con1.Close();
            macro();
        }

        //여신등록
        private void buttonXP14_Click(object sender, EventArgs e)
        {
            string query = "Insert into debt (회원명,이용사,등록일,금액) values (@Name,@Company,@Rdate,@Cost)";
            cmd2 = new OleDbCommand(query, con2);
            cmd2.Parameters.AddWithValue("@Name", richTextBox12.Text);
            cmd2.Parameters.AddWithValue("@Company", comboBox1.SelectedItem);
            cmd2.Parameters.AddWithValue("@Rdate", dateTimePicker1.Value);
            cmd2.Parameters.AddWithValue("@Cost", richTextBox13.Text);


            con2.Open();
            cmd2.ExecuteNonQuery();
            con2.Close();
            debt();
        }

        //계좌삭제
        private void buttonXP5_Click(object sender, EventArgs e)
        {
            try
            {
                string delQuery = "Delete from Sheet1 Where [월렛사용자명]=@Name";
                cmd = new OleDbCommand(delQuery, con);
                cmd.CommandText = delQuery;
                cmd.Connection = con;
                cmd.Parameters.AddWithValue("@Name", richTextBox3.Text);

                MessageBox.Show("계좌삭제완료");


                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                GetAccount();


            }
            catch (Exception)
            {
                MessageBox.Show("오류,다시 시도하세요 ");
            }
        }

        //매크로삭제
        private void buttonXP10_Click(object sender, EventArgs e)
        {
            try
            {
                string delQuery1 = "Delete from macro Where [내용]=@Context";
                cmd1 = new OleDbCommand(delQuery1, con1);
                cmd1.CommandText = delQuery1;
                cmd1.Connection = con1;
                cmd1.Parameters.AddWithValue("@Context", richTextBox7.Text);

                MessageBox.Show("매크로삭제완료");


                con1.Open();
                cmd1.ExecuteNonQuery();
                con1.Close();
                macro();


            }
            catch (Exception)
            {
                MessageBox.Show("오류,다시 시도하세요 ");
            }
        }

        //여신삭제
        private void buttonXP15_Click(object sender, EventArgs e)
        {
            try
            {
                string query = "Delete from debt Where [회원명]=@Name";
                cmd2 = new OleDbCommand(query, con2);
                cmd2.CommandText = query;
                cmd2.Connection = con2;
                cmd2.Parameters.AddWithValue("@Name", richTextBox12.Text);

                MessageBox.Show("여신삭제완료");


                con2.Open();
                cmd2.ExecuteNonQuery();
                con2.Close();
                debt();


            }
            catch (Exception)
            {
                MessageBox.Show("오류,다시 시도하세요 ");
            }
        }





        //계좌수정
        private void buttonXP6_Click(object sender, EventArgs e)
        {
            string query = "Update Sheet1 Set 월렛사용자명=@Name,가상계좌은행=@Bank,가상계좌번호=@Vaccount Where 월렛사용자명 = @Name";
            cmd = new OleDbCommand(query, con);
            cmd.Parameters.AddWithValue("@Name", richTextBox3.Text);
            cmd.Parameters.AddWithValue("@Bank", richTextBox4.Text);
            cmd.Parameters.AddWithValue("@Vaccount", richTextBox5.Text);


            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            GetAccount();
        }

        //매크로수정
        private void buttonXP11_Click(object sender, EventArgs e)
        {
            if (richTextBox9.Text != "" && richTextBox7.Text != "")
            {
                string updatequery = ("UPDATE macro SET [구분=@Division],[내용=@Context] WHERE 내용=@Context");
                cmd1 = new OleDbCommand(updatequery, con1);
                cmd1.Parameters.AddWithValue("@Division", richTextBox9.Text);
                cmd1.Parameters.AddWithValue("@Context", richTextBox7.Text);

                con1.Open();
                cmd1.ExecuteNonQuery();
                con1.Close();
                MessageBox.Show("수정완료");
                macro();
            }
            else
            {
                MessageBox.Show("다시 시도하세요");

            }





        }

        private void 매크로관리TToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form form = new Form2();
            form.Show();
        }

        private void buttonXP7_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox6.Text);
        }



        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView2.SelectedRows != null)
            {

                string a = richTextBox9.Text;
                string b = richTextBox7.Text;

                richTextBox9.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                richTextBox7.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();

                richTextBox8.Text = b;


            }
        }



        private void buttonXP12_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox11.Text);
        }

        private void 정산표입력JToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form form = new Form3();
            form.Show();
        }


        //엑셀파일 가져오기(에이닐)
        

        private void buttonXP9_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox8.Text);
        }

        private void buttonXP17_Click(object sender, EventArgs e)
        {
            richTextBox14.Text = "로얄조회결과:" + richTextBox15.Text + "/승인담당: "+richTextBox16.Text +"/ "+comboBox3.SelectedItem + "/작성자:"+ richTextBox10.Text;

                
         
        }

        private void buttonXP7_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox6.Text);
        }

        private void dataGridView4_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView4.SelectedRows != null)
            {
                
                richTextBox12.Text = dataGridView4.Rows[e.RowIndex].Cells[0].Value.ToString();
                comboBox1.SelectedItem = dataGridView4.Rows[e.RowIndex].Cells[1].Value.ToString();
                richTextBox13.Text = dataGridView4.Rows[e.RowIndex].Cells[3].Value.ToString();
            }
        }

        private void 매크로관리TToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Form form = new Form2();
            form.Show();
        }

        private void 근무자전달사항WToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form form = new Form3();
            form.Show();
        }

        private void buttonXP12_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox11.Text);
        }

        private void buttonXP16_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox14.Text);
        }

        private void 어드민관리AToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }
    }

}


    
