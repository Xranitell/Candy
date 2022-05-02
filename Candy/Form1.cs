using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Candy
{
    public partial class Form1 : Form
    {
        SQLquery sql = new SQLquery();
        public Form1()
        {
            InitializeComponent();
        }
        void UpdateItems()
        {
            OrdersDGV.DataSource = sql.SelectAllOrders();
            ClientsGV.DataSource = sql.SelectAllClients();
            dataGridView1.DataSource = sql.SelectAllSeller();
            ProductDGV.DataSource = sql.SelectAllProducts();
            TypeDGV.DataSource = sql.SelectAllTypes();

            ClientCB.DataSource = sql.CB_ClientSeller("Client");
            ClientCB.DisplayMember = "ФИО";
            ClientCB.ValueMember = "Id";

            SellerCB.DataSource = sql.CB_ClientSeller("Seller");
            SellerCB.DisplayMember = "ФИО";
            SellerCB.ValueMember = "Id";

            ProductCB.DataSource = sql.CB_Product();
            ProductCB.DisplayMember = "ФИО";
            ProductCB.ValueMember = "Id";

            comboBox1.DataSource = sql.CB_Type();
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "Id";
        }
        #region OrederPanel
        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateItems();
        }
        private void AddOrder_Click(object sender, EventArgs e)
        {
            sql.AddOrder(int.Parse(ClientCB.SelectedValue.ToString()), int.Parse(SellerCB.SelectedValue.ToString()), int.Parse(ProductCB.SelectedValue.ToString()),int.Parse(CountUD.Value.ToString()),dateTimePicker1.Value.ToString("yyyy-MM-dd"));
            MessageBox.Show("Заказ добавлен");
            UpdateItems();
        }
        private void EditOrder_Click(object sender, EventArgs e)
        {
            sql.EditOrder(int.Parse(ClientCB.SelectedValue.ToString()), int.Parse(SellerCB.SelectedValue.ToString()), int.Parse(ProductCB.SelectedValue.ToString()), int.Parse(CountUD.Value.ToString()), dateTimePicker1.Value.ToString("yyyy-MM-dd"), int.Parse(OrdersDGV.CurrentRow.Cells[0].Value.ToString()));
            MessageBox.Show("Заказ N"+ int.Parse(OrdersDGV.CurrentRow.Cells[0].Value.ToString()) + " изменен");
            UpdateItems();
        }
        private void DeleteOrder_Click(object sender, EventArgs e)
        {
            if(sql.DeleteFromGridByIndex("Order", int.Parse(OrdersDGV.CurrentRow.Cells[0].Value.ToString())))
            {
                MessageBox.Show("Заказ N" + int.Parse(OrdersDGV.CurrentRow.Cells[0].Value.ToString()) + " удален");
                UpdateItems();
            }
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
        }
        #endregion
        #region ClientPanel
        private void ClientsGV_SelectionChanged(object sender, EventArgs e)
        {
            if(ClientsGV.CurrentRow!=null)
            {
                ClientSurnameText.Text = ClientsGV.CurrentRow.Cells[1].Value.ToString();
                ClientNameText.Text = ClientsGV.CurrentRow.Cells[2].Value.ToString();
                ClientLastNameText.Text = ClientsGV.CurrentRow.Cells[3].Value.ToString();
                maskedTextBox1.Text = ClientsGV.CurrentRow.Cells[4].Value.ToString();
            }
        }
        private void DeleteClient_Click(object sender, EventArgs e)
        {
            if (sql.DeleteFromGridByIndex("Client", int.Parse(ClientsGV.CurrentRow.Cells[0].Value.ToString())))
            {
                MessageBox.Show("Клиент N" + int.Parse(ClientsGV.CurrentRow.Cells[0].Value.ToString()) + " удален");
                UpdateItems();
            }
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
        }
        private void AddClient_Click(object sender, EventArgs e)
        {
            sql.AddClient(ClientSurnameText.Text,ClientNameText.Text,ClientLastNameText.Text,maskedTextBox1.Text);
            MessageBox.Show("Клиент успешно добавлен");
            UpdateItems();
        }
        private void EditClient_Click(object sender, EventArgs e)
        {
            sql.EditClient(ClientSurnameText.Text, ClientNameText.Text, ClientLastNameText.Text, maskedTextBox1.Text, int.Parse(ClientsGV.CurrentRow.Cells[0].Value.ToString()));
            MessageBox.Show("Клиент успешно изменен");
            UpdateItems();
        }

        #endregion

        #region Seller
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                textBox3.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                textBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            sql.AddSeller(textBox3.Text, textBox2.Text, textBox1.Text, maskedTextBox2.Text);
            MessageBox.Show("Продавец успешно добавлен");
            UpdateItems();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (sql.DeleteFromGridByIndex("Seller", int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString())))
            {
                MessageBox.Show("Продавец N" + int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()) + " удален");
                UpdateItems();
            }
            else MessageBox.Show("Невозможно удалить продавца из-за наличия привязаных записей");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            sql.EditSeller(textBox3.Text, textBox2.Text, textBox1.Text, maskedTextBox2.Text, int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()));
            MessageBox.Show("Клиент успешно изменен");
            UpdateItems();
        }

        #endregion
        #region ProductPanel
        private void ProductDGV_SelectionChanged(object sender, EventArgs e)
        {
            if (ProductDGV.CurrentRow != null)
            {
                textBox6.Text = ProductDGV.CurrentRow.Cells[1].Value.ToString();
                comboBox1.Text = ProductDGV.CurrentRow.Cells[2].Value.ToString();
                textBox5.Text = ProductDGV.CurrentRow.Cells[3].Value.ToString();
                dateTimePicker2.Text = ProductDGV.CurrentRow.Cells[4].Value.ToString();
                numericUpDown1.Value = decimal.Parse(ProductDGV.CurrentRow.Cells[5].Value.ToString());
                numericUpDown2.Value = decimal.Parse(ProductDGV.CurrentRow.Cells[6].Value.ToString());
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            sql.AddProduct(textBox6.Text, int.Parse(comboBox1.SelectedValue.ToString()), textBox5.Text, dateTimePicker2.Value.ToString("yyyy-MM-dd"), int.Parse(numericUpDown1.Value.ToString()), numericUpDown2.Value.ToString());
            MessageBox.Show("Товар успешно добавлен");
            UpdateItems();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            sql.EditProduct(textBox6.Text, int.Parse(comboBox1.SelectedValue.ToString()), textBox5.Text, dateTimePicker2.Value.ToString("yyyy-MM-dd"), int.Parse(numericUpDown1.Value.ToString()),numericUpDown2.Value.ToString(),int.Parse(ProductDGV.CurrentRow.Cells[0].Value.ToString()));
            MessageBox.Show("Товар успешно обновлен");
            UpdateItems();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (sql.DeleteFromGridByIndex("Product", int.Parse(ProductDGV.CurrentRow.Cells[0].Value.ToString())))
            {
                MessageBox.Show("Товар N" + int.Parse(ProductDGV.CurrentRow.Cells[0].Value.ToString()) + " удален");
                UpdateItems();
            }
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
        }


        #endregion

        #region TypesPanel


        private void TypeDGV_SelectionChanged(object sender, EventArgs e)
        {
            textBox7.Text = TypeDGV.CurrentRow.Cells[1].Value.ToString();
        }


        #endregion

        private void button8_Click(object sender, EventArgs e)
        {
            sql.AddType(textBox7.Text);
            MessageBox.Show("Вид продукта успешно добавлен");
            UpdateItems();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            sql.EditType(textBox7.Text,int.Parse(TypeDGV.CurrentRow.Cells[0].Value.ToString()));
            MessageBox.Show("Вид продукта успешно обновлен");
            UpdateItems();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (sql.DeleteFromGridByIndex("Type", int.Parse(TypeDGV.CurrentRow.Cells[0].Value.ToString())))
            {
                MessageBox.Show("Вид товара с № " + int.Parse(TypeDGV.CurrentRow.Cells[0].Value.ToString()) + " удален");
                UpdateItems();
            }
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            DataTable All_products = new DataTable();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            if (dateTimePicker3.Text == "" || dateTimePicker4.Text == "")
            {
                ExcelApp.Cells[1, 1] = "Данные об всех отчетах";
                All_products = sql.SelectAllOrders();
            }
            else
            {
                if (dateTimePicker3.Value < dateTimePicker4.Value)
                {
                    All_products = sql.Take_all_records_by_date(dateTimePicker3, dateTimePicker4);
                    ExcelApp.Cells[1, 1] = "Данные об отчетах в период:" + dateTimePicker3.Value + " - " + dateTimePicker4.Value;
                }
                else
                {
                    MessageBox.Show("Неверный промежуток дат"); return;
                }
            }
            ExcelApp.Columns.NumberFormat = "General";
            ExcelApp.Cells[3, 1] = "ID";
            ExcelApp.Columns[1].ColumnWidth = 5;

            ExcelApp.Cells[3, 2] = "Клиент";
            ExcelApp.Columns[2].ColumnWidth = 30;

            ExcelApp.Cells[3, 3] = "Продавец";
            ExcelApp.Columns[3].ColumnWidth = 30;

            ExcelApp.Cells[3, 4] = "Товар";
            ExcelApp.Columns[4].ColumnWidth = 30;

            ExcelApp.Cells[3, 5] = "Количество";
            ExcelApp.Columns[5].ColumnWidth = 10;
            ExcelApp.Columns.NumberFormat = "@";

            ExcelApp.Cells[3, 6] = "Дата заказа";
            ExcelApp.Columns[6].ColumnWidth = 20;
            
            ExcelApp.Cells[3, 7] = "Итоговая стоимость";
            ExcelApp.Columns[6].ColumnWidth = 30;

            for (int i = 0; i < All_products.Rows.Count; i++)
            {
                ExcelApp.Cells[i + 4, 1] = All_products.Rows[i][0].ToString();
                ExcelApp.Cells[i + 4, 2] = All_products.Rows[i][1].ToString();
                ExcelApp.Cells[i + 4, 3] = All_products.Rows[i][2].ToString();
                ExcelApp.Cells[i + 4, 4] = All_products.Rows[i][3].ToString();
                ExcelApp.Cells[i + 4, 5] = All_products.Rows[i][4].ToString();
                ExcelApp.Cells[i + 4, 6] = All_products.Rows[i][5].ToString();
                ExcelApp.Cells[i + 4, 7] = All_products.Rows[i][6].ToString();
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (OrdersDGV.CurrentRow == null) return;
            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document document =
                    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            //add_text
            document.Content.SetRange(0, 0);
            document.Content.Text =
                "ID записи - " + OrdersDGV.CurrentRow.Cells[0].Value.ToString() + Environment.NewLine +
                "Клиент - " + OrdersDGV.CurrentRow.Cells[1].Value.ToString() + Environment.NewLine +
                "Продавец - " + OrdersDGV.CurrentRow.Cells[2].Value.ToString() + Environment.NewLine +
                "Товар - " + OrdersDGV.CurrentRow.Cells[3].Value.ToString() + Environment.NewLine +
                "Количество - " + OrdersDGV.CurrentRow.Cells[4].Value.ToString() + Environment.NewLine +
                "Дата оформления заказа - " + OrdersDGV.CurrentRow.Cells[5].Value.ToString() + Environment.NewLine +
                "Итоговая стоимость - " + OrdersDGV.CurrentRow.Cells[6].Value.ToString() + Environment.NewLine;

            winword.Visible = true;
        }
    }
}
