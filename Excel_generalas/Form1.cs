﻿using System;
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
                CreateTable(); //Tábla létrehozása*/
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

       
        
        private void CreateTable()
        {
            string[] headers = new string[] //tömb amely tartalmazza a tábla fejléceit
           { "Kód",
             "Eladó",
             "Oldal",
             "Kerület",
             "Lift",
             "Szobák száma",
             "Alapterület (m2)",
             "Ár (mFt)",
             "Négyzetméter ár (Ft/m2)"
           };
            for (int i = 0; i < headers.Length; i++)
            {
              xlSheet.Cells[1, i + 1] = headers[i];
            }
           

            object[,] values = new object[Flats.Count, headers.Length];
            int counter = 0;
            foreach (Flat f in Flats)
            {
                values[counter, 0] = f.Code;
                values[counter, 1] = f.Vendor;
                values[counter, 2] = f.Side;
                values[counter, 3] = f.District;
                values[counter, 4] = f.Elevator;
                values[counter, 5] = f.NumberOfRooms;
                values[counter, 6] = f.FloorArea;
                values[counter, 7] = f.Price;
                values[counter, 8] = "";
                counter++;
            }
            xlSheet.get_Range(
             GetCell(2, 1),
             GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;

            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true; // fejléc betűstílusa félkövér
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // fejléc szövegének középre helyezése függőlegesen
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // fejléc szövegének középre helyezése vízszintesen
            headerRange.EntireColumn.AutoFit(); //akkorára állítja a cellát amekkora az oszlop szélessége
            headerRange.RowHeight = 40; // sormagasság
            headerRange.Interior.Color = Color.LightBlue; //fejléc háttérszíne kék
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick); // fejlécnek vastag keret

        

        }



        private void LoadData()
        {
            Flats = context.Flats.ToList(); //Adattábla memóriába való másolása
        }

        
        
        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }
        
    
    
    
    
    }
}
