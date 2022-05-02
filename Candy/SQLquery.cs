using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Candy
{
    public class SQLquery
    {
        string adress = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=CandyBD;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        #region OrderDGV
        public DataTable SelectAllOrders()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select [Order].Id, Concat(Client.Surname,' ',Client.Name,' ',Client.LastName) as Клиент, Concat(Seller.Surname,' ',Seller.Name,' ',Seller.LastName) as Продавец, CONCAT(Type.Name,' =',Product.Name,'= ', Product.CostByOne, N'р') as 'Товар и стоимость за единицу' ,[Order].CountOfProduct as 'Количество в единицах', [Order].Date as 'Дата заказа', [Order].Summary as 'Итоговая стоимость' From [Order] left join Product on Product.Id = [Order].idProduct inner join Type on Type.Id = Product.IdType left join Client on Client.Id = [Order].idClient left join Seller on Seller.Id = [Order].idSeller";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable CB_ClientSeller(string table)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Id, concat(Surname,' ',Name,' ',LastName) as ФИО from "+ table;
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable CB_Product()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Product.Id, concat(Type.Name,' ',Product.Name) as ФИО from Product left join Type on Type.Id = Product.IdType";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }

        public double SelectCost(int idProduct)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select * from Product where Id = " + idProduct;
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return double.Parse(dt.Rows[0][6].ToString());
        }
        public void AddOrder(int ClientID, int SellerID,int ProductID,int Count,string date)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into [Order](idClient,idSeller,idProduct,CountOfProduct,Date,Summary) Values(N'" + ClientID + "',N'" + SellerID + "',N'" + ProductID + "', N'" + Count + "', N'" + date + "',N'" + (SelectCost(ProductID) * Count).ToString().Replace(",",".") + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditOrder(int ClientID, int SellerID, int ProductID, int Count, string date, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update [Order] set IdClient = N'" + ClientID + "', IdSeller = N'" + SellerID + "' , IdProduct = N'" + ProductID + "' , CountOfProduct = N'" + Count + "' , Date = N'" + date + "', Summary = N'" + SelectCost(ProductID) * Count + "' where Id = " + id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        #endregion

        #region ClientsDGV
        public DataTable SelectAllClients()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Id,Surname as 'Фамилия', Name as 'Имя', LastName as 'Отчество', Phone as 'Номер телефона' From Client";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public void AddClient(string Surname, string Name, string LastName, string Phone)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into [Client](Surname,Name,LastName,Phone) Values(N'" + Surname + "',N'" + Name + "',N'" + LastName + "', N'" + Phone + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditClient(string Surname, string Name, string LastName, string Phone, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update Client set Surname = N'" + Surname + "', Name = N'" + Name + "' , LastName = N'" + LastName + "', Phone = N'" + Phone + "' where Id = " + id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        #endregion


        #region SellerDGV

        public DataTable SelectAllSeller()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select id, Surname as 'Фамилия', Name as 'Имя', LastName as 'Отчество', Phone as 'Номер телефона' From Seller";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public void AddSeller(string Surname, string Name, string LastName, string Phone)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into [Seller](Surname,Name,LastName,Phone) Values(N'" + Surname + "',N'" + Name + "',N'" + LastName + "', N'" + Phone + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditSeller(string Surname, string Name, string LastName, string Phone, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update Seller set Surname = N'" + Surname + "', Name = N'" + Name + "' , LastName = N'" + LastName + "', Phone = N'" + Phone + "' where Id = " + id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        #endregion

        #region ProductsDVG
        public DataTable SelectAllProducts()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Product.Id, Product.Name as 'Название',Type.Name as 'Тип',Product.Structure as 'Состав',Product.DateOfCreate as 'Дата создания',Product.LifeTime as 'Срок хранения в днях',Product.CostByOne as 'Стоимость за единицу' From Product left join Type on Type.Id = Product.IdType";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable CB_Type()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Id,Name from Type";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }

        public void AddProduct(string Name, int IdType, string Structure, string DateOfCreate,int LifeTime, string CostByOne)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into [Product](Name,IdType, Structure, DateOfCreate, LifeTime, CostByOne) Values(N'" + Name + "',N'" + IdType + "',N'" + Structure + "', N'" + DateOfCreate + "', N'" + LifeTime + "', N'" + CostByOne.Replace(",", ".") + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditProduct(string Name, int IdType, string Structure, string DateOfCreate, int LifeTime, string CostByOne,int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update Product set Name = N'" + Name + "', IdType = N'" + IdType + "' , Structure = N'" + Structure + "', DateOfCreate = N'" + DateOfCreate + "', LifeTime = N'" + LifeTime + "', CostByOne = '" +  CostByOne.Replace(",",".") + "' where Id = " + id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        #endregion

        #region TypePanel
        public DataTable SelectAllTypes()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Id, Name as Название From Type";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public void AddType(string Type)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into [Type](Name) Values(N'" + Type + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditType(string Type, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update Type set Name = N'" + Type + "' where Id = " + id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        #endregion

        public DataTable Take_all_records_by_date(DateTimePicker firstDate, DateTimePicker lastDate)
        {
            SqlConnection con = new SqlConnection(); ;
            SqlCommand cmd = new SqlCommand(); ;
            DataTable table = new DataTable();   //объявляю адаптер, таблицу, конектор, команду
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.ConnectionString = adress;    //Присваиваю коннектору строку подключения к бд  SELECT T1.key, T2.name FROM table1 T1 JOIN table2 T2 ON T1.name = T2.key;
            cmd.Connection = con;                       //комманде задается коннектор
            con.Open();
            cmd.CommandText = "Select [Order].Id, Concat(Client.Surname,' ',Client.Name,' ',Client.LastName) as Клиент, Concat(Seller.Surname,' ',Seller.Name,' ',Seller.LastName) as Продавец, CONCAT(Type.Name,' =',Product.Name,'= ', Product.CostByOne, N'р') as 'Товар и стоимость за единицу' ,[Order].CountOfProduct as 'Количество в единицах', [Order].Date as 'Дата заказа', [Order].Summary as 'Итоговая стоимость' From [Order] left join Product on Product.Id = [Order].idProduct inner join Type on Type.Id = Product.IdType left join Client on Client.Id = [Order].idClient left join Seller on Seller.Id = [Order].idSeller where [Order].Date between '" + firstDate.Value.ToString("yyyy-MM-dd") + "' and '" + lastDate.Value.ToString("yyyy-MM-dd") + "'";
            adapt.SelectCommand = cmd;
            adapt.Fill(table);   //заполняю таблицу полученными данными
            return table;
        }
        #region ForAll
        public bool DeleteFromGridByIndex(string Table, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            con.Open();
            cmd.CommandText = "Delete from [" + Table + "] where Id = " + id;
            cmd.Connection = con;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                return false;
            }
            con.Close();
            return true;
        }
        #endregion

    }
    

    
}
