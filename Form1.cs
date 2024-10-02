using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.Text.RegularExpressions;
using System.Threading;
using System.Reflection;

namespace EBAY
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            populate();
            
            //this.FormClosed += MyClosedHandler;

          //  enable_connection();
        }

        

        public class category_name
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }
        void populate()
        {
            var dataSource = new List<category_name>();
            MySqlConnection connection = new MySqlConnection("SERVER=127.0.0.1; DATABASE=ebay; UID=root; PASSWORD=;Connect Timeout=300");
            connection.Open();

            MySqlCommand cmd = new MySqlCommand("SELECT categories_id, categories_name FROM categories_description where language_id=2 and categories_id!=0 order by categories_id", connection);// 

            MySqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                int category_id = (int)rdr["categories_id"];
                string category_name = (string)rdr["categories_name"];
                dataSource.Add(new category_name() { Name = category_name, Value = category_id.ToString() });
            }


            //Setup data binding
            this.comboBox1.DataSource = dataSource;
            this.comboBox1.DisplayMember = "Name";
            this.comboBox1.ValueMember = "Value";
            this.comboBox1.SelectedIndex = -1;

            // make it readonly
            this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;


            Thread.Sleep(3000);
            
        }
        private void button1_Click_1(object sender, EventArgs e)//generate links
        {
            listBox1.Items.Clear();

            if (comboBox1.SelectedIndex < 0) { MessageBox.Show("Please select a category first !"); }
            else
            {

                for (int i = 1; i <= Int32.Parse(pages.Text); i++)
                {
                    listBox1.Items.Add(url.Text + "&_pgn=" + i);
                }
            }
        }

             


        //functions
        static string listimg(List<string> img)
        {
            string data = "";
            foreach (string imgs in img)
            {
                data += imgs + ";";
            }
            return data;
        }
        static string clean(string data)//clean item description
        {
            string[] res = new string[] { "<TD>", "&nbsp;", "<TR>", "</TD>", "</TR>", "<TD width=\"50%\">", "<TD class=attrLabels>", "<SPAN>", "</SPAN>", "<TABLE role=presentation cellSpacing=0 cellPadding=0 width=\"100%\">", "<TBODY>", "</TBODY>", "</TABLE>", "</DIV>", "<DIV class=section>", "<DIV aria-live=polite>" };

            foreach (string c in res)
            {
                data = data.Replace(c, " ");
                data = data.Trim();
            }
            int st = data.IndexOf("<H2");
            int en = data.IndexOf("</H2>") + 5;
           if(st>-1) data = data.Remove(st, en - st);
            data = data.Trim();
            int i = 0;
            while ((i = data.IndexOf("<!--", i)) != -1)
            {
                en = data.IndexOf("-->", i) + 3;
                data = data.Remove(i, en - i);
            }
            i = 0;
            while ((i = data.IndexOf("\r\n", i)) != -1)
            {
                en = data.IndexOf("\r\n", i) + 3;
                data = data.Remove(i, en - i);
            }
            i = 0;
            while ((i = data.IndexOf("  ", i)) != -1)
            {
                data = data.Replace("  ", " ");
                i++;
            }
            return data;
        }
       

        private void button4_Click(object sender, EventArgs e)
        {

            string category_name = comboBox1.GetItemText(comboBox1.SelectedItem);
            int category_id = Int32.Parse(comboBox1.GetItemText(comboBox1.SelectedValue));

            SqlConnection connection = new SqlConnection("Data Source=JAMES\\SQLEXPRESS;Initial Catalog=ebay1;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            connection.Open();


            for (int scan = 0; scan < listBox1.Items.Count; scan++)
            {
                {
                    webBrowser1.Navigate(listBox1.Items[scan].ToString(), null, null, "User-Agent:affiliate_partner");

                    while (webBrowser1.ReadyState != WebBrowserReadyState.Complete) Application.DoEvents();

                    try {
                        if (webBrowser1.Document != null)
                        {
                            HtmlElementCollection elem_producta = webBrowser1.Document.GetElementsByTagName("a");


                            foreach (HtmlElement link in elem_producta)
                            {
                                if (link.GetAttribute("className") == "vip")
                                {
                                    string data = link.OuterHtml;
                                    int st = data.IndexOf("href=") + 6;
                                    int en = data.IndexOf("\"", st);

                                    data = data.Substring(st, en - st);
                                    en = data.IndexOf("?");
                                    data = data.Substring(0, en);

                                    listBox2.Items.Add(data);
                                    Application.DoEvents();
                                    listBox2.TopIndex = listBox2.Items.Count - 1;
                                    label3.Text = "Results :" + listBox2.Items.Count.ToString();

                                    //  Console.WriteLine(data);



                                    SqlCommand cmd = new SqlCommand("SELECT count(*) FROM pages where files=@Files", connection);
                                    SqlParameter param = new SqlParameter();
                                    param.ParameterName = "@Files";
                                    param.Value = data;
                                    cmd.Parameters.Add(param);
                                    int count = (int)cmd.ExecuteScalar();
                                    if (count > 0) { }
                                    else
                                    {
                                        SqlCommand cmd1 = new SqlCommand("INSERT INTO pages (files,category_id,category_name, scan) values (@Files, @Category_id, @Category_name, @Scan)", connection);

                                        param = new SqlParameter();
                                        param.ParameterName = "@Category_id";
                                        param.Value = category_id;
                                        cmd1.Parameters.Add(param);

                                        param = new SqlParameter();
                                        param.ParameterName = "@Files";
                                        param.Value = data;
                                        cmd1.Parameters.Add(param);

                                        param = new SqlParameter();
                                        param.ParameterName = "@Scan";
                                        param.Value = 1;
                                        cmd1.Parameters.Add(param);

                                        param = new SqlParameter();
                                        param.ParameterName = "@Category_name";
                                        param.Value = category_name;
                                        cmd1.Parameters.Add(param);


                                        cmd1.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                    }catch(Exception ex) { }
                }
            }
            MessageBox.Show("creating links over");

        }

        int Product_id_num()
        {
            int id = 0;
            MySqlConnection connectionmy = new MySqlConnection("SERVER=127.0.0.1; DATABASE=ebay; UID=root; PASSWORD=;Connect Timeout=300");
            connectionmy.Open();
            MySqlCommand cmd1 = new MySqlCommand("SELECT products_id FROM products order by products_id DESC", connectionmy);
            MySqlDataReader rdr1 = cmd1.ExecuteReader();
            if (rdr1.Read())
            {
                id = (int)rdr1["products_id"];
            }
            rdr1.Close();
            return id + 1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                listView1.Items[i].Selected = true;


                string itemname =           listView1.Items[i].SubItems[0].Text;
                string ratingcount =        listView1.Items[i].SubItems[1].Text;
                string ratingvalue =        listView1.Items[i].SubItems[2].Text;
                string itemprice =          listView1.Items[i].SubItems[3].Text;
                string pricecurrency =      listView1.Items[i].SubItems[4].Text;      
                string itemcondition=       listView1.Items[i].SubItems[5].Text;
                string availableatorfrom =  listView1.Items[i].SubItems[6].Text;
                string author =             listView1.Items[i].SubItems[7].Text;
                string authorid =           listView1.Items[i].SubItems[8].Text;
                string authorvalue =        listView1.Items[i].SubItems[9].Text;
                string availablefor =       listView1.Items[i].SubItems[10].Text;
                string itemid =             listView1.Items[i].SubItems[11].Text;
                string itemqty =            listView1.Items[i].SubItems[12].Text;
                string itemshipping =       listView1.Items[i].SubItems[13].Text;
                string itemattribute =      listView1.Items[i].SubItems[14].Text;   
                string itemimage =          listView1.Items[i].SubItems[15].Text;



                int category_id = Int32.Parse(comboBox1.GetItemText(comboBox1.SelectedValue));


                string[] _img = itemimage.Split(';');
                int j = 0;
                string thumbimg = thumb_image(_img[0], itemname);

                foreach (string img in _img)
                {
                    j++;
                    if (img != thumbimg)
                    {
                       save_images(img, itemid, itemname, j, Product_id_num());
                    }
                }

                if (itemid == "" || itemname == "" || itemimage=="") { }
                else
                {
               
                MySqlConnection myconnection = new MySqlConnection("SERVER=127.0.0.1; DATABASE=ebay; UID=root; PASSWORD=;Connect Timeout=300");
                    myconnection.Open();

                    int p_id= Product_id_num();

                    MySqlCommand mycmd = new MySqlCommand("INSERT INTO products_to_categories (products_id, categories_id) values(@Product_id,@Category_id)", myconnection);
                    MySqlParameter myparam = new MySqlParameter();
                    myparam.ParameterName = "@Product_id";
                    myparam.Value = p_id;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Category_id";
                    myparam.Value = category_id;
                    mycmd.Parameters.Add(myparam);

                    mycmd.ExecuteNonQuery();
                                       
                    mycmd = new MySqlCommand("INSERT INTO saler (saler_id, saler_name, saler_rating, products_id) values(@Saler_id, @Saler_name, @Saler_rating, @Products_id)", myconnection);
                                        
                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Saler_id";
                    myparam.Value = authorid;
                    mycmd.Parameters.Add(myparam);
                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Saler_name";
                    myparam.Value = author;
                    mycmd.Parameters.Add(myparam);
                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Saler_rating";
                    myparam.Value = authorvalue;
                    mycmd.Parameters.Add(myparam);
                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_id";
                    myparam.Value = p_id;

                    mycmd.Parameters.Add(myparam);

                    mycmd.ExecuteNonQuery();


                    mycmd = new MySqlCommand("INSERT INTO products (products_quantity,products_price, products_image,products_shipping_cost,products_price_currency,products_ebay_id,products_status) values(@Products_quantity,@Products_price,@Products_image,@Products_shipping_cost,@Products_price_currency,@Products_ebay_id,@Products_status)", myconnection);
                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_quantity";
                    myparam.Value = itemqty;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_price";
                    myparam.Value = itemprice;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_shipping_cost";
                    myparam.Value = itemshipping;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_image";
                    myparam.Value = thumbimg;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_price_currency";
                    myparam.Value = pricecurrency;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_ebay_id";
                    myparam.Value = itemid;
                    mycmd.Parameters.Add(myparam);

                    myparam = new MySqlParameter();
                    myparam.ParameterName = "@Products_status";
                    myparam.Value = 1;
                    mycmd.Parameters.Add(myparam);

                    mycmd.ExecuteNonQuery();

                    MySqlCommand mycmd1 = new MySqlCommand("INSERT INTO products_description (products_name,language_id,products_description,products_condition,products_availableatforfrom,products_availablefor) values (@Products_name, @Language_id,@Products_description,@Products_condition,@Products_availableatforfrom,@Products_availablefor)", myconnection);

                    MySqlParameter myparam1 = new MySqlParameter();

                    
                    myparam1.ParameterName = "@Products_name";
                    myparam1.Value = itemname;
                    mycmd1.Parameters.Add(myparam1);

                    myparam1 = new MySqlParameter();
                    myparam1.ParameterName = "@Language_id";
                    myparam1.Value = 2;
                    mycmd1.Parameters.Add(myparam1);

                    myparam1 = new MySqlParameter();
                    myparam1.ParameterName = "@Products_description";
                    myparam1.Value = itemattribute;
                    mycmd1.Parameters.Add(myparam1);

                    myparam1 = new MySqlParameter();
                    myparam1.ParameterName = "@Products_availableatforfrom";
                    myparam1.Value = availableatorfrom;
                    mycmd1.Parameters.Add(myparam1);

                    myparam1 = new MySqlParameter();
                    myparam1.ParameterName = "@Products_availablefor";
                    myparam1.Value = availablefor;
                    mycmd1.Parameters.Add(myparam1);

                    myparam1 = new MySqlParameter();
                    myparam1.ParameterName = "@Products_condition";
                    myparam1.Value = itemcondition;
                    mycmd1.Parameters.Add(myparam1);

                    mycmd1.ExecuteNonQuery();

                   
                     myconnection.Close();

            


                }
            }

            MessageBox.Show("saved !");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string category_name = comboBox1.GetItemText(comboBox1.SelectedItem);
            int category_id = Int32.Parse(comboBox1.GetItemText(comboBox1.SelectedValue));

            SqlConnection connection = new SqlConnection("Data Source=JAMES\\SQLEXPRESS;Initial Catalog=ebay;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            connection.Open();

            for (int i = 0; i < listBox2.Items.Count; i++)
            {
               string html = string.Empty;
                string uri = listBox2.Items[i].ToString();
                string files = uri;
                // Console.WriteLine(uri);
                /*
               string saveLocation = "C:\\Ebay\\temp\\" + files;


               HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);

               using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
               using (Stream stream = response.GetResponseStream())
               using (StreamReader reader = new StreamReader(stream))
               {
                   html = reader.ReadToEnd();
               }

               FileStream fs = new FileStream(saveLocation, FileMode.Create);
               BinaryWriter bw = new BinaryWriter(fs);
               try
               {
                   bw.Write(html);
               }
               finally
               {
                   fs.Close();
                   bw.Close();
               }
               Thread.Sleep(1000);

               Application.DoEvents();

   */

                int j = 0;
                SqlCommand cmd = new SqlCommand("SELECT count(*) FROM pages where files=@Files", connection);
                SqlParameter param = new SqlParameter();
                param.ParameterName = "@Files";
                param.Value = files;
                cmd.Parameters.Add(param);
                int count = (int)cmd.ExecuteScalar();
                if (count > 0) { }
                else
                {
                    SqlCommand cmd1 = new SqlCommand("INSERT INTO pages (files,category_id,category_name, scan) values (@Files, @Category_id, @Category_name, @Scan)", connection);

                    param = new SqlParameter();
                    param.ParameterName = "@Category_id";
                    param.Value = category_id;
                    cmd1.Parameters.Add(param);

                    param = new SqlParameter();
                    param.ParameterName = "@Files";
                    param.Value = files;
                    cmd1.Parameters.Add(param);

                    param = new SqlParameter();
                    param.ParameterName = "@Scan";
                    param.Value = 1;
                    cmd1.Parameters.Add(param);

                    param = new SqlParameter();
                    param.ParameterName = "@Category_name";
                    param.Value = category_name;
                    cmd1.Parameters.Add(param);

                    j++;
                    cmd1.ExecuteNonQuery();
                }
                label3.Text = j.ToString() + " files added";
            }

            MessageBox.Show("Saving OVER");
        }

        void disable_connection()
        {
            System.Diagnostics.Process.Start("ipconfig", "/release"); //For disabling internet
            Thread.Sleep(4000);
        }
        void enable_connection()
        {
            System.Diagnostics.Process.Start("ipconfig", "/renew"); //For enabling internet
            Thread.Sleep(3000);
        }

      
        void check()
        {

            int category_id = Int32.Parse(comboBox1.GetItemText(comboBox1.SelectedValue));
            string category_name = comboBox1.Text;

            SqlConnection connection1 = new SqlConnection("Data Source=JAMES\\SQLEXPRESS;Initial Catalog=ebay1;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            connection1.Open();

            //select undone links

            SqlCommand cmd1 = new SqlCommand("SELECT files FROM pages where scan=@Scan and category_id=@category_id", connection1);
            SqlParameter param = new SqlParameter();
            param.ParameterName = "@Scan";
            param.Value = 1;
            cmd1.Parameters.Add(param);

            param = new SqlParameter();
            param.ParameterName = "@Category_id";
            param.Value = category_id;
            cmd1.Parameters.Add(param);

            //  param = new SqlParameter();
            // param.ParameterName = "@Id";
            // param.Value = 17;
            // cmd1.Parameters.Add(param);

            SqlDataReader reader = cmd1.ExecuteReader();

            while (reader.Read())
            {
                string files = (string)reader["files"];

                listBox2.Items.Add(files);

            }
            label3.Text = "Results: " + listBox2.Items.Count.ToString();

        }




        private void button9_Click(object sender, EventArgs e)
        {
            check();
        }

        protected void MyClosedHandler(object sender, EventArgs e)
        {
            enable_connection();
        }

        static string thumb_image(string imgurl, string itemname)
        {
            string uri = imgurl.Replace("64", "800");
            itemname = img_clean_url(itemname);
            string file = itemname.ToLower().Replace(" ", "-").Replace("/", "_")  + ".jpg";

            return file;
        }

        static void save_images(string imgurl, string itemid, string itemname, int j, int products_id)
        {
            string file;
            string html = string.Empty;
            string uri = imgurl.Replace("64", "800");

            if (uri != "" && itemid != "" && itemname != "" && products_id > 0)
            {

                itemname = img_clean_url(itemname);

                if (j == 1) { 
                file = itemname.Trim().ToLower().Replace(" ", "-").Replace("/", "_") +  ".jpg";
            }
            else
            {

                file = itemname.Trim().ToLower().Replace(" ", "-").Replace("/", "_") + "_" + j + ".jpg";
            }
                                

                MySqlConnection myconnection = new MySqlConnection("SERVER=127.0.0.1; DATABASE=ebay; UID=root; PASSWORD=;Connect Timeout=300");
                myconnection.Open();
                MySqlCommand mycmd2 = new MySqlCommand("INSERT INTO products_images (products_id,image,sort_order) values (@Products_id, @Products_image,@Sort_order)", myconnection);
                MySqlParameter myparam2 = new MySqlParameter();
                myparam2.ParameterName = "@Products_id";
                myparam2.Value = products_id;
                mycmd2.Parameters.Add(myparam2);
                myparam2 = new MySqlParameter();
                myparam2.ParameterName = "@Products_image";
                myparam2.Value = file;
                mycmd2.Parameters.Add(myparam2);
                myparam2 = new MySqlParameter();
                myparam2.ParameterName = "@Sort_order";
                myparam2.Value = j;
                mycmd2.Parameters.Add(myparam2);

                mycmd2.ExecuteNonQuery();
                myconnection.Close();

                Directory.CreateDirectory("c:\\ebay\\item_images\\" + itemid);

               // Console.WriteLine(file);

                string saveLocation = "C:\\Ebay\\item_images\\" + itemid + "\\" + file;

                using (WebClient client = new WebClient())
                {
                    try
                    {
                        client.DownloadFile(new Uri(uri), saveLocation);
                    }
                    catch (Exception ex) { }
                }
 
              //  Thread.Sleep(1000);

                Application.DoEvents();
            }
        }
        static string img_clean_url(string url)
        {
            url = url.Replace("é", "e");
            url = url.Replace("è", "e");
            url = url.Replace("ê", "e");
            url = url.Replace("à", "a");
            url = url.Replace("ï", "i");

            url = Regex.Replace(url, "[^0-9a-zA-Z]+", "-");
            url = url.Replace("-pl", "");
            return url;

       
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string category_name = comboBox1.GetItemText(comboBox1.SelectedItem);
            int category_id = Int32.Parse(comboBox1.GetItemText(comboBox1.SelectedValue));

            listView1.Items.Clear();

            SqlConnection connection1 = new SqlConnection("Data Source=JAMES\\SQLEXPRESS;Initial Catalog=ebay1;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            connection1.Open();

            //select undone links

            SqlCommand cmd1 = new SqlCommand("SELECT itemname,ratingcount,ratingvalue,itemprice,pricecurrency,availableatorfrom,pricecurrency,author,authorid,authorvalue,availablefor,itemid,itemqty,itemshipping,itemattribute,itemimages,itemcondition FROM ebay where category_id=@category_id", connection1);
            SqlParameter param = new SqlParameter();

            param.ParameterName = "@Category_id";
            param.Value = category_id;
            cmd1.Parameters.Add(param);

            SqlDataReader reader = cmd1.ExecuteReader();

            while (reader.Read())
            {
                string itemname = (string)reader["itemname"];
                string ratingvalue = (string)reader["ratingvalue"];
                string ratingcount = (string)reader["ratingcount"];
                string itemprice = (string)reader["itemprice"];
                string pricecurrency = (string)reader["pricecurrency"];
                string itemcondition= (string)reader["itemcondition"];
                string availableatorfrom = (string)reader["availableatorfrom"];
                string author = (string)reader["author"];
                string authorid = (string)reader["authorvalue"];
                string authorvalue = (string)reader["authorid"];
                string availablefor = (string)reader["availablefor"];
                string itemid = (string)reader["itemid"];
                string itemqty = (string)reader["itemqty"];
                string itemshipping = (string)reader["itemshipping"];
                string itemattribute = (string)reader["itemattribute"];
                string itemimages = (string)reader["itemimages"];



                ListViewItem item = listView1.Items.Add(itemname);
                item.SubItems.Add(ratingvalue);
                item.SubItems.Add(ratingcount);
                item.SubItems.Add(itemprice);
                item.SubItems.Add(pricecurrency);
                item.SubItems.Add(itemcondition);
                item.SubItems.Add(availableatorfrom);
                item.SubItems.Add(author);
                item.SubItems.Add(authorid);
                item.SubItems.Add(authorvalue);
                item.SubItems.Add(availablefor);
                item.SubItems.Add(itemid);
                item.SubItems.Add(itemqty);
                item.SubItems.Add(itemshipping);
                item.SubItems.Add(itemattribute);
                item.SubItems.Add(itemimages);

                listView1.EnsureVisible(listView1.Items.Count - 1);

                Application.DoEvents();
            }
            label4.Text = "Results : " + listView1.Items.Count.ToString();
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        // navigate WebBrowser to the list of urls in a loop
        static async Task<object> DoWorkAsync(object[] args)//SCAN
        {
            DateTime TempsDeDepart = DateTime.Now;


            SqlConnection connection = new SqlConnection("Data Source=JAMES\\SQLEXPRESS;Initial Catalog=ebay1;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            connection.Open();


            string itemname = "";
            string ratingvalue = "";
            string ratingcount = "";
            string itemprice = "";
            string pricecurrency = "";
            string itemcondition = "";
            string availableatorfrom = "";
            string author = "";
            string authorid = "";
            string authorvalue = "";
            string availablefor = "";
            string itemid = "";
            string itemqty = "";
            string itemshipping = "";
            string itemattribute = "";
            string itemimage = "";

            int category_id = 82;// Int32.Parse(args[0].ToString());

            Console.WriteLine("Start working.");

            using (var wb = new WebBrowser())
            {
                wb.ScriptErrorsSuppressed = true;

                TaskCompletionSource<bool> tcs = null;
                WebBrowserDocumentCompletedEventHandler documentCompletedHandler = (s, e) =>
                tcs.TrySetResult(true);

              //  MessageBox.Show(args.Length.ToString());

                // navigate to each URL in the list
                foreach(var url in args) 
                {
                   // string url = args[w].ToString();

                    MessageBox.Show(url.ToString());
                  
                    Console.WriteLine(url.ToString());
                    tcs = new TaskCompletionSource<bool>();
                    wb.DocumentCompleted += documentCompletedHandler;
                    try
                    {
                        wb.Navigate(url.ToString());
                        // await for DocumentCompleted
                        await tcs.Task;
                    }
                    finally
                    {
                        wb.DocumentCompleted -= documentCompletedHandler;
                    }
                    // the DOM is ready

                    //  Console.WriteLine(wb.Document.Body.OuterHtml);


                    try
                    {
                        HtmlElement elem_producth1 = wb.Document.GetElementById("vi-lkhdr-itmTitl");

                        if (elem_producth1 != null)
                        {
                            string title = elem_producth1.InnerText.ToLower();

                            itemname = title;
                        }


                        int st = 0;
                        int en = 0;

                        HtmlElementCollection elem_productspan = wb.Document.GetElementsByTagName("span");

                        if (elem_productspan.Count > 0)
                        {
                            for (int ii = 0; ii < elem_productspan.Count; ii++)
                            {
                                string result = elem_productspan[ii].OuterHtml.ToLower();

                                if (result.IndexOf("span itemprop=\"ratingvalue\" content") > 0)
                                {
                                    st = result.IndexOf("content=") + 9;
                                    en = result.IndexOf(">", st) - 1;
                                    ratingvalue = result.Substring(st, en - st);
                                }
                                if (result.IndexOf("span itemprop=\"reviewcount\" content") > 0)
                                {
                                    st = result.IndexOf("content=") + 9;
                                    en = result.IndexOf(">", st) - 1;
                                    ratingcount = result.Substring(st, en - st);
                                }

                                if (result.IndexOf("itemprop=\"pricecurrency\"") > 0)
                                {
                                    st = result.IndexOf("content=") + 9;
                                    en = result.IndexOf(">", st) - 1;
                                    pricecurrency = result.Substring(st, en - st);
                                }

                                if (result.IndexOf("itemprop=\"availableatorfrom\"") > 0)//div
                                {
                                    st = result.IndexOf("itemprop=\"availableatorfrom\"") + 29;
                                    en = result.IndexOf("<", st) - 1;
                                    availableatorfrom = result.Substring(st, en - st);
                                }

                                if (result.IndexOf("itemprop=\"areaserved\"") > 0)//div
                                {
                                    st = result.IndexOf("itemprop=\"areaserved\"") + 22;
                                    en = result.IndexOf("<", st);
                                    availablefor = result.Substring(st, en - st);
                                }
                            }
                        }

                        HtmlElementCollection elem_producta = wb.Document.GetElementsByTagName("span");
                        foreach (HtmlElement link in elem_producta)
                        {
                            if (link.GetAttribute("className") == "mbg-nw")
                            {
                                author = link.InnerText;
                            }
                        }
                        HtmlElement elem_productdiv3 = wb.Document.GetElementById("prcIsum");
                        if (elem_productdiv3 != null) itemprice = elem_productdiv3.InnerText.Replace(",", ".").Replace("EUR", "").Replace("/pièce", "").Replace("GBP", "").Replace("USD", "");

                        HtmlElementCollection elemdiv = wb.Document.GetElementsByTagName("div");
                        foreach (HtmlElement link in elemdiv)
                        {
                            if (link.GetAttribute("className") == "mbg")
                            {
                                string data = link.InnerHtml;
                                st = data.IndexOf("_trksid=") + 8;
                                en = data.IndexOf("\"", st);
                                authorid = data.Substring(st, en - st);
                            }
                        }
                        if (authorid == "") { authorid = "particulier"; }



                        HtmlElement elem_productb = wb.Document.GetElementById("qtySubTxt");
                        if (elem_productb != null)
                        {
                            itemqty = elem_productb.InnerText;
                        }
                        else
                        {
                            itemqty = "1";
                        }

                        if (itemqty == "") { }
                        HtmlElement elem_productbc = wb.Document.GetElementById("fshippingCost");
                        if (elem_productbc != null) itemshipping = elem_productbc.InnerText;

                        // caracteristique de l'objet
                        HtmlElement elem_productdiv1 = wb.Document.GetElementById("viTabs_0_is");
                        if (elem_productdiv1 != null) itemattribute = elem_productdiv1.InnerHtml;

                        HtmlElement elem_hidden1 = wb.Document.GetElementById("vi-cond-addl-info");
                        if (elem_hidden1 != null) itemattribute = itemattribute.Replace(elem_hidden1.OuterHtml, "");

                        HtmlElement elem_hidden2 = wb.Document.GetElementById("hiddenContent");
                        if (elem_hidden2 != null) itemattribute = itemattribute.Replace(elem_hidden2.OuterHtml, "");

                        HtmlElement elem_hidden3 = wb.Document.GetElementById("readFull");
                        if (elem_hidden3 != null) itemattribute = itemattribute.Replace(elem_hidden3.OuterHtml, "");

                        string[] res = new string[] { "<TD>", "&nbsp;", "<TR>", "</TD>", "</TR>", "<TD width=\"50%\">", "<TD class=attrLabels>", "<SPAN>", "</SPAN>", "<TABLE role=presentation cellSpacing=0 cellPadding=0 width=\"100%\">", "<TBODY>", "</TBODY>", "</TABLE>", "</DIV>", "<DIV class=section>", "<DIV aria-live=polite>" };

                        foreach (string c in res)
                        {
                            itemattribute = itemattribute.Replace(c, " ");
                            itemattribute = itemattribute.Trim();
                        }
                        st = itemattribute.IndexOf("<H2");
                        en = itemattribute.IndexOf("</H2>") + 5;
                        if (st > -1) itemattribute = itemattribute.Remove(st, en - st);
                        itemattribute = itemattribute.Trim();
                        int i = 0;
                        while ((i = itemattribute.IndexOf("<!--", i)) != -1)
                        {
                            en = itemattribute.IndexOf("-->", i) + 3;
                            itemattribute = itemattribute.Remove(i, en - i);
                        }
                        i = 0;
                        while ((i = itemattribute.IndexOf("\r\n", i)) != -1)
                        {
                            en = itemattribute.IndexOf("\r\n", i) + 3;
                            itemattribute = itemattribute.Remove(i, en - i);
                        }
                        i = 0;
                        while ((i = itemattribute.IndexOf("  ", i)) != -1)
                        {
                            itemattribute = itemattribute.Replace("  ", " ");
                            i++;
                        }




                        HtmlElement elem_productdiv = wb.Document.GetElementById("si-fb");
                        if (elem_productdiv != null) authorvalue = elem_productdiv.InnerText.Substring(0, elem_productdiv.InnerText.IndexOf("%") + 1);

                        HtmlElement elem_productcond = wb.Document.GetElementById("vi-itm-cond");
                        if (elem_productcond != null) itemcondition = elem_productcond.InnerText;

                        HtmlElement elem_productnumber = wb.Document.GetElementById("descItemNumber");
                        if (elem_productnumber != null) itemid = elem_productnumber.InnerText;

                        //images
                        HtmlElementCollection elem_productimg = wb.Document.GetElementsByTagName("td");
                        foreach (HtmlElement link in elem_productimg)
                        {
                            if (link.GetAttribute("className") == "tdThumb")
                            {
                                string res1 = link.OuterHtml;
                                st = res1.IndexOf(" src=") + 6;
                                en = res1.IndexOf("\"", st);
                                res1 = res1.Substring(st, en - st);
                                itemimage += res1 + ";";

                            }
                        }

                        if (itemimage == "")
                        {
                            HtmlElement thumb = wb.Document.GetElementById("icImg");
                            if (thumb.InnerHtml != null)
                            {
                                string data = thumb.InnerHtml;
                                st = data.IndexOf("src=");
                                en = data.IndexOf("\"", st);
                                if (st > en) itemimage = data.Substring(st, en - st);
                            }
                        }

                        itemimage = itemimage.Replace("é", "e");
                        itemimage = itemimage.Replace("è", "e");
                        itemimage = itemimage.Replace("ê", "e");
                        itemimage = itemimage.Replace("à", "a");
                        itemimage = itemimage.Replace("ï", "i");

                        itemimage = Regex.Replace(itemimage, "[^0-9a-zA-Z]+", "-");
                        itemimage = itemimage.Replace("-pl", "");

                        itemname = itemname.Replace("é", "e");
                        itemname = itemname.Replace("è", "e");
                        itemname = itemname.Replace("ê", "e");
                        itemname = itemname.Replace("à", "a");
                        itemname = itemname.Replace("ï", "i");

                        itemname = Regex.Replace(itemname, "[^0-9a-zA-Z]+", "-");

                        itemimage = itemimage.Replace("64", "800");

                        itemimage = itemname.ToLower().Replace(" ", "-").Replace("/", "_") + ".jpg";



                        /*
                         *  {
            string uri = imgurl.Replace("64", "800");
            itemname = img_clean_url(itemname);
            string file = itemname.ToLower().Replace(" ", "-").Replace("/", "_")  + ".jpg";

            return file;
            */

                        Console.WriteLine(itemname);// + " " + ratingvalue + " " + ratingcount + " " + itemprice + " " + pricecurrency + " " + itemcondition + " " + availableatorfrom + " " + author + " " + authorid + " " + authorvalue + " " + availablefor + " " + itemid + " " + itemqty + " " + itemshipping + " " + itemattribute + " " + itemimage);

                        DateTime nowe = DateTime.Now;
                        TimeSpan timee = nowe.Subtract(TempsDeDepart);
                        Console.WriteLine("Elapse time  : " + timee);

                        Application.DoEvents();

                        if (itemimage != "")
                        {
                                 
                            SqlCommand cmd = new SqlCommand("SELECT count(*) FROM ebay where itemid=@Itemid", connection);
                            SqlParameter param = new SqlParameter();
                            param.ParameterName = "@Itemid";
                            param.Value = itemid;
                            cmd.Parameters.Add(param);

                            int count = (int)cmd.ExecuteScalar();

                            if (count > 0) { } else{
                                //insert
                                cmd = new SqlCommand("INSERT INTO ebay (itemid,itemname,itemprice,pricecurrency,ratingvalue,ratingcount,availableatorfrom,author,authorid,authorvalue," +
                                    "availablefor,itemqty,itemshipping,itemattribute, itemimages,itemcondition, category_id) values (@Itemid, @Itemname,@Itemprice," +
                                    "@Pricecurrency, @Ratingvalue, @Ratingcount, @Availableatorfrom,@Author,@Authorvalue,@Authorid,@Availablefor,@Itemqty,@Itemshipping,@Itemattribute, @Itemimages, @itemcondition,@Category_id)", connection);
                                param = new SqlParameter();
                                param.ParameterName = "@Itemid";
                                param.Value = itemid;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Itemname";
                                param.Value = itemname;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Itemprice";
                                param.Value = itemprice;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Pricecurrency";
                                param.Value = pricecurrency;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Ratingvalue";
                                param.Value = ratingvalue;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Ratingcount";
                                param.Value = ratingcount;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Availableatorfrom";
                                param.Value = availableatorfrom;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Availablefor";
                                param.Value = availablefor;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Itemqty";
                                param.Value = itemqty;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Itemshipping";
                                param.Value = itemshipping;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Itemattribute";
                                param.Value = itemattribute;
                                cmd.Parameters.Add(param);


                                param = new SqlParameter();
                                param.ParameterName = "@Itemcondition";
                                param.Value = itemcondition;
                                cmd.Parameters.Add(param);


                                param = new SqlParameter();
                                param.ParameterName = "@Itemimages";
                                param.Value = itemimage;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Author";
                                param.Value = author;
                                cmd.Parameters.Add(param);


                                param = new SqlParameter();
                                param.ParameterName = "@Authorid";
                                param.Value = authorid;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Authorvalue";
                                param.Value = authorvalue;
                                cmd.Parameters.Add(param);

                                param = new SqlParameter();
                                param.ParameterName = "@Category_id";
                                param.Value = category_id;
                                cmd.Parameters.Add(param);

                                cmd.ExecuteNonQuery();

                            }
                        }
                    }
                    catch (Exception ex) { }
                }
            }

            Console.WriteLine("End working.");
            return null;
        }
        private void button3_Click(object sender, EventArgs e)//scan page downloaded
        {
          //  Main();
        }


        static void Main(string[] args, ComboBox comboBox1, ListBox listBox2) {

            string category_id = comboBox1.GetItemText(comboBox1.SelectedValue);

            for (int a = 0; a < listBox2.Items.Count; a++)
            {
             
                string uris = "";
                for (int j = a; j < a + 5; j++)
                {
                    uris += "\"" + listBox2.Items[j] + "\",";
                }
                string uri = uris.Substring(0, uris.Length - 1);
                

              //  string uri = listBox2.Items[a].ToString();

                try
                {
                    // download each page and dump the content

                    var task = MessageLoopWorker.Run(DoWorkAsync, uri);


                    task.Wait();
                    Console.WriteLine("DoWorkAsync completed.");
                }

                catch (Exception ex)
                {
                    Console.WriteLine("DoWorkAsync failed: " + ex.Message);
                }
               a += 5;
            }
           // check();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }

    public static class MessageLoopWorker
    {
        public static async Task<object> Run(Func<object[], Task<object>> worker, params object[] args)
        {
            var tcs = new TaskCompletionSource<object>();

            var thread = new Thread(() =>
            {
                EventHandler idleHandler = null;

                idleHandler = async (s, e) =>
                {
                    // handle Application.Idle just once
                    Application.Idle -= idleHandler;

                    // return to the message loop
                    await Task.Yield();

                    // and continue asynchronously
                    // propogate the result or exception
                    try
                    {
                        var result = await worker(args);
                        tcs.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }

                    // signal to exit the message loop
                    // Application.Run will exit at this point
                    Application.ExitThread();
                };

                // handle Application.Idle just once
                // to make sure we're inside the message loop
                // and SynchronizationContext has been correctly installed
                Application.Idle += idleHandler;
                Application.Run();
            });

            // set STA model for the new thread
            thread.SetApartmentState(ApartmentState.STA);

            // start the thread and await for the task
            thread.Start();
            try
            {
                return await tcs.Task;
            }
            finally
            {
                thread.Join();
            }
        }
    }

}
