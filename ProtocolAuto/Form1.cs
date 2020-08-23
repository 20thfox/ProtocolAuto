/*
* Заказчик (label) - costumer
		 (textBox) - costumerBox

* Объект    (label) - objct 
		  (textBox) - objctBox

* Пусконаладочная орг-ция (label) - agency 
				        (textBox) - agencyBox

* Присоединение: (label) - objAdd 
	           (textBox) - objAddBox

* Номер протокола: (textBox) - protNumBox

* Температура (label) - temp
	        (textBox) - tempBox

* Атмосферное давление (label) - press
		(textBox) - pressBox

* Влажность (label) - wet
		  (textBox) - wetBox

* Испытания: (label) - test
		   (textBox) - testBox

* Дата испытания (label) - dateTest
	           (textBox) - dateTestBox

* Дата регистрации (label) - dateReg
	             (textBox) - dateRegBox

* Результаты проверил (label) - audit
		(textBox) - auditBox

* Испытания произвели (label) - testPers
				   (comboBox) - testPersBox1
				   (comboBox) - testPersBox2
				   (comboBox) - testPersBox3

* ФИО (testBox) - fioBox1
	  (testBox) - fioBox2
	  (testBox) - fioBox3

tabControlPanel
mainForm
*/
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

namespace ProtocolAuto
{
    public partial class mainForm : Form //Тест синхронизации11
    {
        private Word.Application wordapp; //глобальное определение Word.Application
        private Word.Document worddocument;
        public mainForm()
        {
            InitializeComponent();
        }

        private void creat_Click(object sender, EventArgs e)//нажатие кнопки Создать
        {
            if(emptyTest()==false)//проверка на заполненость полей
            {
                wordapp = new Word.Application();
                wordapp.Visible = true;
                genTemplate();
            }
        }
        private void genTemplate()
        {
            try
            {
                Object template = Environment.CurrentDirectory + @"\Templates\Example.docx";//получает путь к exe + путь к файлу, Example заменить на переменную
                Object newTemplate = false;
                Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;


                worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                genFormat();


               // Save();

            }
            catch (Exception)
            {
                /*
                wordapp.Quit(ref falseObj, ref missingObj, ref missingObj);
                worddocument = null;
                wordapp = null;
                genFaultActive();
                */
            }
        }
        private void genFormat()
        {
            //Объявление всякой хрени
            Object findText;
            Object replaceText;
            //
            //Замена объекта и присоединения
            //
            Word.Table table1 = worddocument.Tables[1]; //Обращение к таблице по индексу 1
            table1.Cell(1, 4).Range.InsertAfter(costumerBox.Text); //вставка значения поля в ячейку таблицы
            table1.Cell(2, 4).Range.InsertAfter(objctBox.Text);
            table1.Cell(3, 4).Range.InsertAfter(agencyBox.Text);
            table1.Cell(4, 4).Range.InsertAfter(objAddBox.Text);
            //
            //Замента номера протокола, температуры, давления и влаги
            //
            findText = "п00-0-0-0000";
            replaceText = protNumBox.Text + "-" + "ЗАМЕНИТЬ" + "-" + DateTime.Now.Year.ToString();
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@test";
            replaceText = testBox.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@Temp";
            replaceText = tempBox.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@Pres";
            replaceText = pressBox.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);
            findText = "@Vlag";
            replaceText = wetBox.Text;
            wordapp.Selection.Find.Execute(ref findText, ReplaceWith: ref replaceText);
            wordapp.Selection.Collapse(0);

            //
            //замена нижнего колонтитула
            //
            foreach (Word.Section sec in worddocument.Sections)
            {
                var range = sec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                Word.Table table3 = range.Tables[1];
                table3.Cell(1, 1).Range.InsertAfter(protNumBox.Text + "-" + "ЗАМЕНИТЬ" + "-" + DateTime.Now.Year.ToString());
            }

            //
            //Испытания произвели, фамилии даты и прочее
            //

            var countTabl = worddocument.Tables.Count;
            Word.Table lastTable = worddocument.Tables[countTabl];
            lastTable.Cell(1, 2).Range.InsertAfter(testPersBox1.Text);
            lastTable.Cell(2, 2).Range.InsertAfter(testPersBox2.Text);
            lastTable.Cell(3, 2).Range.InsertAfter(testPersBox3.Text);
            lastTable.Cell(4, 2).Range.InsertAfter(auditBox.Text);
            lastTable.Cell(1, 3).Range.InsertAfter(fioBox1.Text);
            lastTable.Cell(2, 3).Range.InsertAfter(fioBox2.Text);
            lastTable.Cell(3, 3).Range.InsertAfter(fioBox3.Text);
            lastTable.Cell(4, 3).Range.InsertAfter(fioBox4.Text);
            lastTable.Cell(5, 2).Range.InsertAfter(dateRegBox.Text);
            lastTable.Cell(6, 2).Range.InsertAfter(dateTestBox.Text);
        }
        private bool emptyTest()
        {
            var listTextBox = new List<TextBox> 
                {costumerBox, //список текстбоксов
                objctBox,
                agencyBox,
                objAddBox,
                protNumBox,
                tempBox,
                pressBox,
                wetBox,
                testBox,
                dateTestBox,
                dateRegBox,
                auditBox,
                fioBox4,
                fioBox3,
                fioBox2,
                fioBox1};
            var listComboBox = new List<ComboBox> {testPersBox1, testPersBox2, testPersBox3};//список комбобоксов
            bool empty = false;//переменная для проверки заполнености при true - не все ячейки заполнены

        //
        //проверка текстбоксов на заполненость
        //
            foreach (var txtB in listTextBox)
            {
                if (txtB.Text.Length == 0)
                {
                    txtB.BackColor = Color.MistyRose;
                    empty = true;
                }
                else
                {
                    txtB.BackColor = Color.White;
                }
            }
            //
            //проверка комбобоксов на заполненость
            //
            foreach (var txtB in listComboBox)
            {
                if (txtB.Text.Length == 0)
                {
                    txtB.BackColor = Color.MistyRose;
                    empty = true;
                }
                else
                {
                    txtB.BackColor = Color.White;
                }
            }
            if (empty == true)//вывод сообщения предупреждения
            {
                MessageBox.Show("Заполните все выделенные поля!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return empty;//возврат заполнености
        }
    }
}
