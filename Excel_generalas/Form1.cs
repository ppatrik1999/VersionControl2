using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.Entity.Migrations.Model;

namespace Excel_generalas
{
    public partial class Form1 : Form
    {
        RealEstateEntities context = new RealEstateEntities(); // ORM objektum példányosítása
        List<Flat> Flats; // Flat tipusú elemekből álló listára mutató referencia

        Excel.Application xlApp; // A microsoft excel alkalmazás
        Excel.Workbook xlWB; // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül


        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();
        }

        private void CreateExcel()
        {
            try
            {
                xlApp = new Excel.Application(); //Excel elindítása és az applikáció objektum betöltése
                xlWB = xlApp.Workbooks.Add(Missing.Value); // Új munkafüzet
                xlSheet = xlWB.ActiveSheet; // Új munkalap
                /*CreateTable(); //Tábla létrehozása*/
                xlApp.Visible = true; //Control átadása a felhasználónak
                xlApp.UserControl = true; //Control átadása a felhasználónak
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");
                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
            
        }

       /* private void CreateTable()
        {
            throw new NotImplementedException();
        }*/

        private void LoadData()
        {
            Flats = context.Flats.ToList(); //Adattábla memóriába való másolása
        }
    }
}
