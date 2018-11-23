using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace favorite_folder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public List<String> filelist;
        public string defaultpath= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private void Form1_Load(object sender, EventArgs e)
        {
            ImageList iconList = new ImageList();
            iconList.ImageSize = new Size(60, 60);
            iconList.Images.Add(
                Image.FromFile(Application.StartupPath + @"\..\..\Resources\Excel.png"));
            iconList.Images.Add(
                Image.FromFile(Application.StartupPath + @"\..\..\Resources\word.png"));
            iconList.Images.Add(
                Image.FromFile(Application.StartupPath + @"\..\..\Resources\ppt.png"));
            iconList.Images.Add(
                Image.FromFile(Application.StartupPath + @"\..\..\Resources\pdf.png"));
            iconList.Images.Add(
                Image.FromFile(Application.StartupPath + @"\..\..\Resources\file.png"));
            iconList.Images.Add(
                Image.FromFile(Application.StartupPath + @"\..\..\Resources\Folder.png"));
            listView_main.Items.Clear();
            listView_main.LargeImageList = iconList;

            listView_main.Columns.Add("File", 50);
            listView_main.View = View.LargeIcon;

            listView_main.Font = new Font(FontFamily.GenericSansSerif,
            14.0F, FontStyle.Bold);

            filelist = new List<string>();
            InitialList();
        }
        //================= form move ==================================================
        bool formMove = false;//窗体是否移动
        Point formPoint;//记录窗体的位置

        private void form_MouseDown(object sender, MouseEventArgs e)
        {
            formPoint = new Point();
            int xOffset;
            int yOffset;
            if (e.Button == MouseButtons.Left)
            {
                xOffset = -e.X - SystemInformation.FrameBorderSize.Width;
                yOffset = -e.Y - SystemInformation.CaptionHeight - SystemInformation.FrameBorderSize.Height;
                formPoint = new Point(xOffset, yOffset);
                formMove = true;//开始移动
            }
        }

        private void form_MouseMove(object sender, MouseEventArgs e)
        {
            if (formMove == true)
            {
                Point mousePos = Control.MousePosition;
                mousePos.Offset(formPoint.X, formPoint.Y);
                Location = mousePos;
            }
        }

        private void form_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)//按下的是鼠标左键
            {
                formMove = false;//停止移动
            }
        }

        private void button_exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void listView_main_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Items.Clear();
                ImageList mageList1 = new ImageList();
                mageList1.Images.Add(
                    Image.FromFile(Application.StartupPath + @"\..\..\Resources\excute.png"));
                mageList1.Images.Add(
                    Image.FromFile(Application.StartupPath + @"\..\..\Resources\add.png"));
                mageList1.Images.Add(
                    Image.FromFile(Application.StartupPath + @"\..\..\Resources\open.png"));
                mageList1.Images.Add(
                    Image.FromFile(Application.StartupPath + @"\..\..\Resources\delete.png"));

                contextMenuStrip1.Items.Add("Excute", mageList1.Images[0]);
                contextMenuStrip1.Items.Add("Add File", mageList1.Images[1]);
                contextMenuStrip1.Items.Add("Add Folder", mageList1.Images[2]);
                contextMenuStrip1.Items.Add("Remove", mageList1.Images[3]);

                contextMenuStrip1.Show(MousePosition);
            }
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

            switch (e.ClickedItem.Text)
            {
                case "Excute":
                    if (listView_main.SelectedItems.Count>0)
                    {
                        
                        listView_main.Cursor = Cursors.AppStarting;
                        Thread excute_thread = new Thread( new ParameterizedThreadStart(excute_program));
                        excute_thread.Start(filelist[listView_main.SelectedItems[0].Index]);
                        listView_main.Cursor = Cursors.Default;
                    }
                    break;
                case "Add File":
                    Addnewfile();
                    break;
                case "Add Folder":
                    Addnewfolder();
                    break;
                case "Remove":
                    Removefile();
                    break;
            }
            
        }

        private void Addnewfolder()
        {
            folderBrowserDialog1.Tag = "選擇資料夾";
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filelist.Add(folderBrowserDialog1.SelectedPath);
            }
            refreshfilelist();
        }
        private void InitialList()
        {
             String readfilelist="";
             if (!File.Exists(defaultpath + "//favoritelist.txt"))
             {
                 var myFile = File.Create(defaultpath + "//favoritelist.txt"); ;
                 myFile.Close();
                 
                 
             }
             StreamReader file = new StreamReader(defaultpath+"//favoritelist.txt");
             readfilelist = file.ReadToEnd();
             file.Close();
             String[] filestring = readfilelist.Split('\n');
             foreach (String filename in filestring)
             {
                 if (filename.Length>5)
                 {
                    filelist.Add(filename);
                 }
             }
             refreshfilelist();
        }
        private void Addnewfile()
        {
            openFileDialog1.Tag = "選擇檔案";
            
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filelist.Add(openFileDialog1.FileName);
            }
            refreshfilelist();
        }
        private void Removefile()
        {
            int index = listView_main.SelectedItems[0].Index;
            if(index>=0)
                filelist.RemoveAt(index);
            refreshfilelist();
        }
        private void refreshfilelist()
        {
            listView_main.Items.Clear();
            foreach (String filename in filelist)
            {
                FileInfo file=new FileInfo(filename);
                switch (file.Extension)
                {
                    case ".xlsx":
                    case ".xls":
                        listView_main.Items.Add(System.IO.Path.GetFileName(filename), 0);
                        break;
                    case ".doc":
                    case ".docx":
                        listView_main.Items.Add(System.IO.Path.GetFileName(filename), 1);
                        break;
                    case ".ppt":
                    case ".pptx":
                        listView_main.Items.Add(System.IO.Path.GetFileName(filename), 2);
                        break;
                    case ".pdf":
                        listView_main.Items.Add(System.IO.Path.GetFileName(filename), 3);
                        break;
                    case "":
                        listView_main.Items.Add(System.IO.Path.GetFileName(filename), 5);
                        break;
                    default:
                        listView_main.Items.Add(System.IO.Path.GetFileName(filename), 4);
                        break;
                }
                
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            String savefilelist="";
            foreach(String filename in filelist)
            {
                savefilelist += filename + "\n";
            }
 /*           Properties.Settings.Default.Filename_string = savefilelist;
            Properties.Settings.Default.Save();*/
            System.IO.StreamWriter listfilewriter =
    new System.IO.StreamWriter(defaultpath + "//favoritelist.txt");
            listfilewriter.Write(savefilelist);
            listfilewriter.Close();
        }

        private void listView_main_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Point localPoint = listView_main.PointToClient(MousePosition);
            var item = listView_main.GetItemAt(localPoint.X, localPoint.Y);
            int index = item.Index;
            if (index >= 0)
            {
                listView_main.Cursor = Cursors.AppStarting;
                Thread excute_thread = new Thread(new ParameterizedThreadStart(excute_program));
                excute_thread.Start(filelist[listView_main.SelectedItems[0].Index]);
                
            }
        }

        private void listView_main_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.All);
        }
        private int Dropindex;
        private void listView_main_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.AllowedEffect;
            var pos = listView_main.PointToClient(new Point(e.X, e.Y));
            var hit = listView_main.HitTest(pos);
            Dropindex = hit.Item.Index;
        }

        private void listView_main_DragDrop(object sender, DragEventArgs e)
        {
            var pos = listView_main.PointToClient(new Point(e.X, e.Y));
            var hit = listView_main.HitTest(pos);
            int index=hit.Item.Index;

            String temp = filelist[index];
            filelist[index] = filelist[Dropindex];
            filelist[Dropindex] = temp;

            refreshfilelist();
        }
        //bool isinthread = false;
        private void excute_program(object location)
        {
            Process.Start((string)location);
            updateCursor();

        }
        private void updateCursor()
        {
            Func<int> del;
            del = delegate()
            {
                listView_main.Cursor = Cursors.Default;
                return 0;
            };
            Invoke(del);
        }
        
    }
}
