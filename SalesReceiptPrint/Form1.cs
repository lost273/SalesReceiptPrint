using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SalesReceiptPrint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            dataGridView1.Columns.Add("1", "Наименование товара (услуги)");
            dataGridView1.Columns.Add("2", "кол-во");
            dataGridView1.Columns.Add("3", "цена(руб.коп)");
            dataGridView1.Columns.Add("4", "сумма(руб.коп)");
            dataGridView1.RowCount = 9;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView1.Height = dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible)+6;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.AllowUserToResizeColumns = false;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[0]).MaxInputLength = 33;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[1]).MaxInputLength = 4;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[2]).MaxInputLength = 7;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[3]).MaxInputLength = 7;
            


            // Ячейки с суммой ставим readonly
            for (int i = 0; i < 8; i++) dataGridView1.Rows[i].Cells[3].ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
            
            //printPreviewDialog1.ShowDialog();

        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            StringFormat str = new StringFormat();
            Font drawFont = new Font("Verdana", 10);
            Font drawFontBig = new Font("Verdana", 14);
            Font drawFontBoldCol = new Font("Verdana", 10, FontStyle.Bold);
            Font drawFontBold = new Font("Verdana", 12, FontStyle.Bold | FontStyle.Italic); 
            str.Alignment = StringAlignment.Near;
            str.LineAlignment = StringAlignment.Center;
            str.Trimming = StringTrimming.EllipsisCharacter;

            int contWidth = 430; // ширина ячейки наименования
            int width2 = 65; // ширина 2й ячейки
            int width3 = 130; // ширина 3й ячейки
            int width4 = 135; // ширина 4й ячейки
            int widthСoordinate = 15; // координата ширины

            int height = 25; // высота строки
            int heightСoordinate = 15; // координата высоты


            // Рисуем ООО "Сантиум", УНП 290719859, г.Брест, ул.Кирова, д.50
            e.Graphics.DrawString("ООО \"Сантиум\", УНП 290719859, г.Брест, ул.Кирова, д.50", drawFontBig, Brushes.Black, widthСoordinate, heightСoordinate);
            heightСoordinate = 45;

            // Чек и номер
            e.Graphics.DrawString("ТОВАРНЫЙ ЧЕК № " + textBox1.Text + " от " + dateTimePicker1.Text
                , drawFontBold, Brushes.Black, widthСoordinate, heightСoordinate);
            heightСoordinate = 75;

            // Рисуем названия колонок
            e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, contWidth, height);
            e.Graphics.DrawString("Наименование товара (услуги)", drawFontBoldCol, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
            widthСoordinate += contWidth;
            e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, width2, height);
            e.Graphics.DrawString("кол-во", drawFontBoldCol, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
            widthСoordinate += width2;
            e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, width3, height);
            e.Graphics.DrawString("цена(руб.коп)", drawFontBoldCol, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
            widthСoordinate += width3;
            e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, width4, height);
            e.Graphics.DrawString("сумма(руб.коп)", drawFontBoldCol, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
            heightСoordinate += height;

            // Рисуем остальную таблицу
            for (int i = 0; i < 8; i++)
            {
                widthСoordinate = 15;
                e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, contWidth, height);
                e.Graphics.DrawString((dataGridView1.Rows[i].Cells[0].Value ?? string.Empty).ToString(), drawFont, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
                widthСoordinate += contWidth;
                e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, width2, height);
                e.Graphics.DrawString((dataGridView1.Rows[i].Cells[1].Value ?? string.Empty).ToString(), drawFont, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
                widthСoordinate += width2;
                e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, width3, height);
                e.Graphics.DrawString((dataGridView1.Rows[i].Cells[2].Value ?? string.Empty).ToString(), drawFont, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
                widthСoordinate += width3;
                e.Graphics.DrawRectangle(Pens.Black, widthСoordinate, heightСoordinate, width4, height);
                e.Graphics.DrawString((dataGridView1.Rows[i].Cells[3].Value ?? string.Empty).ToString(), drawFont, Brushes.Black, widthСoordinate+3, heightСoordinate+3);
                heightСoordinate += height;
            }

            // Рисуем Итого
            widthСoordinate = 15;
            e.Graphics.DrawString("ИТОГО: " + label3.Text, drawFontBold, Brushes.Black, widthСoordinate, heightСoordinate);
            heightСoordinate += height*2;

            // Рисуем Подпись продавца
            widthСoordinate = contWidth;
            e.Graphics.DrawString("Подпись продавца _______________________", drawFont, Brushes.Black, widthСoordinate, heightСoordinate);

            

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            double totalSum = 0;
            string fPartSum, sPartSum;

            // Если изменения происходят в 1 или 2 столбце считаем 3 столбец
            if ((e.ColumnIndex == 1) || (e.ColumnIndex == 2))
            {
                // Проверяем ячейки на отсутсвие информации
                if ((dataGridView1.Rows[e.RowIndex].Cells[1].Value == null) || (dataGridView1.Rows[e.RowIndex].Cells[2].Value == null))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[3].Value = null;
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells[3].Value =
                   (Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[2].Value)).ToString("N2");
                    
                    // Приводим ячейки к денежному формату
                    if (dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString().IndexOf(',') == -1)
                        dataGridView1.Rows[e.RowIndex].Cells[2].Value = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString() + ",00";
                    if (dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString().IndexOf(',') == -1)
                        dataGridView1.Rows[e.RowIndex].Cells[3].Value = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString() + ",00";
                }
                
                // Формируем строку ИТОГО
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[3].Value != null) 
                        totalSum = totalSum + Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                }

                if (totalSum != 0)
                {
                    fPartSum = totalSum.ToString("N2");
                    sPartSum = totalSum.ToString("N2");
                    fPartSum = fPartSum.Remove(fPartSum.IndexOf(','));
                    sPartSum = sPartSum.Remove(0,sPartSum.Length-2);
                    label3.Text = totalSum.ToString("N2") + " (" + NumberInWords(Convert.ToInt32(fPartSum)) + " руб. " + sPartSum + " коп." + " )";
                }
                else label3.Text = "0";
            }            
            
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox tb = (TextBox)e.Control;
            tb.KeyPress -= tb_KeyPress;
            tb.KeyPress += new KeyPressEventHandler(tb_KeyPress);
        }

        // Обработка нашатия клавиши в ячейке
        void tb_KeyPress(object sender, KeyPressEventArgs e)
        {

            string vlCell = ((TextBox)sender).Text;

            // Проверяем ввод денежного числа
            if (dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].IsInEditMode == true)

                if ((e.KeyChar == '.') || (e.KeyChar == ','))
                {
                    e.KeyChar = ',';
                    if (vlCell.IndexOf(",") != -1) e.Handled = true;

                }
                else
                 if ((vlCell.Length - vlCell.IndexOf(',') > 2) && (vlCell.IndexOf(',') != -1) || ((e.KeyChar < '0') || (e.KeyChar > '9')))
                {
                    if (e.KeyChar == Convert.ToChar(Keys.Back)) return;

                    e.Handled = true;
                };

        }

        // Сумма прописью
        string NumberInWords(int number)
        {
            int[] array_int = new int[4];
            string[,] array_string = new string[4, 3] {{" миллиард", " миллиарда", " миллиардов"},
                {" миллион", " миллиона", " миллионов"},
                {" тысяча", " тысячи", " тысяч"},
                {"", "", ""}};
            array_int[0] = (number - (number % 1000000000)) / 1000000000;
            array_int[1] = ((number % 1000000000) - (number % 1000000)) / 1000000;
            array_int[2] = ((number % 1000000) - (number % 1000)) / 1000;
            array_int[3] = number % 1000;
            string result = "";
            for (int i = 0; i < 4; i++)
            {
                if (array_int[i] != 0)
                {
                    if (((array_int[i] - (array_int[i] % 100)) / 100) != 0)
                    {
                        switch (((array_int[i] - (array_int[i] % 100)) / 100))
                        {
                            case 1: result += " сто"; break;
                            case 2: result += " двести"; break;
                            case 3: result += " триста"; break;
                            case 4: result += " четыреста"; break;
                            case 5: result += " пятьсот"; break;
                            case 6: result += " шестьсот"; break;
                            case 7: result += " семьсот"; break;
                            case 8: result += " восемьсот"; break;
                            case 9: result += " девятьсот"; break;
                        }
                    }
                    if (((array_int[i] % 100) - ((array_int[i] % 100) % 10)) / 10 != 1)
                    {
                        switch (((array_int[i] % 100) - ((array_int[i] % 100) % 10)) / 10)
                        {
                            case 2: result += " двадцать"; break;
                            case 3: result += " тридцать"; break;
                            case 4: result += " сорок"; break;
                            case 5: result += " пятьдесят"; break;
                            case 6: result += " шестьдесят"; break;
                            case 7: result += " семьдесят"; break;
                            case 8: result += " восемьдесят"; break;
                            case 9: result += " девяносто"; break;
                        }
                    }
                    if (array_int[i] % 100 >= 10 && array_int[i] % 100 <= 19)
                    {
                        switch (array_int[i] % 100)
                        {
                            case 10: result += " десять"; break;
                            case 11: result += " одиннадцать"; break;
                            case 12: result += " двенадцать"; break;
                            case 13: result += " тринадцать"; break;
                            case 14: result += " четырнадцать"; break;
                            case 15: result += " пятнадцать"; break;
                            case 16: result += " шестнадцать"; break;
                            case 17: result += " семнадцать"; break;
                            case 18: result += " восемннадцать"; break;
                            case 19: result += " девятнадцать"; break;
                        }
                    } else
                    switch (array_int[i] % 10)
                    {
                        case 1: if (i == 2) result += " одна"; else result += " один"; break;
                        case 2: if (i == 2) result += " две"; else result += " два"; break;
                        case 3: result += " три"; break;
                        case 4: result += " четыре"; break;
                        case 5: result += " пять"; break;
                        case 6: result += " шесть"; break;
                        case 7: result += " семь"; break;
                        case 8: result += " восемь"; break;
                        case 9: result += " девять"; break;
                    } 
                }
                if (array_int[i] % 100 >= 10 && array_int[i] % 100 <= 19) result += " " + array_string[i, 2] + " ";

                else switch (array_int[i] % 10)
                    {
                        case 1: result += " " + array_string[i, 0] + " "; break;
                        case 2:
                        case 3:
                        case 4: result += " " + array_string[i, 1] + " "; break;
                        case 5:
                        case 6:
                        case 7:
                        case 8:
                        case 9: result += " " + array_string[i, 2] + " "; break;
                    }
            }
            return result;
        }
    }
}
