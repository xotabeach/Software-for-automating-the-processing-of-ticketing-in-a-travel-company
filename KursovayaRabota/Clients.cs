using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace KursovayaRabota
{

    class Clients
    {
        TextInfo ti = CultureInfo.CurrentCulture.TextInfo;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        //private Excel.Range excelcells;
        //protected string surname;
        protected string name;
        protected string phonetic;
        protected string home_adress;
        protected string phone_number;
        protected string[,] ClientsBase;
        protected List<string> baseClient = new List<string>();
        public string[,] ClientBase
        {
            get { return ClientsBase; }
        }
        public void WriteList()
        {
            string text = "";

            int h = 0;
            int k = 0;
            baseClient = new List<string>();
            //StreamReader read_cl_base = new StreamReader("ClientsBase.txt");
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            excelsheets = excelAppworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            Excel.Range xlRange = excelworksheet.UsedRange;
            ClientsBase = new string[xlRange.Count / 5, 5];
            for (int i = 2; i <= xlRange.Count / 5; i++)
            {
                for (char j = 'A'; j <= 'E'; j++)
                {

                    //Console.WriteLine(xlRange.Count);
                    var excelcells = excelworksheet.get_Range($"{j}{i}", Type.Missing);
                    string sStr = Convert.ToString(excelcells.Value2);

                    text += sStr + " ";
                    ClientsBase[h, k] = sStr;
                    k++;

                }
                k = 0;
                h++;
                baseClient.Add(text);
                text = "";
            }
            excelapp.Quit();
            //

            /*В данном конструкторе произошло заполнение массива базы клиентов, где
             * каждая строка содержала информацию об определенном человеке, а номер ряда содержал определенный тип данных
             * (1-Фамилия
             *  2-Имя
             *  3-Отчество
             *  4-Адрес
             *  5-Номер телефона)
             */

        }
        public void PrintInfo()
        {
            string info = " ";
            for (int i = 0; i < ClientsBase.GetLength(0); i++)
            {
                for (int j = 0; j < ClientsBase.GetLength(1); j++)
                {
                    info += $"  {ClientsBase[i, j]}  ";

                }
                Console.WriteLine($"\n Клиент№{i + 1}" + info);
                info = "";
            }
        }
        public void AddNewClient(string surname, string name, string phonetic, string homeadress, string phonenumber)
        {

            string newClient = $"{surname} {name} {phonetic} {homeadress} {phonenumber}";
            string[] client = newClient.Split(" ");
            int i = 0;
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            var excelsheets = excelAppworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            Excel.Range xlRange = excelworksheet.UsedRange;
            for (char j = 'A'; j <= 'E'; j++)
            {
                var excelcells = excelworksheet.get_Range($"{j}{(xlRange.Count / 5) + 1}");
                excelcells.Value2 = client[i];
                i++;
            }
            var excelappworkbooks = excelapp.Workbooks;
            var excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Save();
            excelapp.Quit();
            WriteList();
        }
        public void TakeDataClient(string surname)
        {
            char space = ' ';
            char newspace = '_';
            string answertocheck = "";
            int found = 0;
            bool answer = true;
            Console.Write($"Хорошо, введите вашу фамилию: {surname}\n");

            Console.Write("\nХорошо, теперь введите ваше имя: ");
            name = Console.ReadLine();
            found = name.IndexOf(":");
            name = name.Substring(found + 1);
            name = ti.ToLower(name);
            name = ti.ToTitleCase(name);

            Console.Write("\nВведите ваше отчество: ");
            phonetic = Console.ReadLine();
            found = phonetic.IndexOf(":");
            phonetic = phonetic.Substring(found + 1);
            phonetic = ti.ToLower(phonetic);
            phonetic = ti.ToTitleCase(phonetic);

            Console.Write("\nВведите ваш адрес: ");
            home_adress = Console.ReadLine();
            found = home_adress.IndexOf(":");
            home_adress = home_adress.Substring(found + 1);
            home_adress = home_adress.Replace(space, newspace);

            Console.Write("\nВведите ваш номер телефона: ");
            phone_number = Console.ReadLine();
            found = phone_number.IndexOf(":");
            phone_number = phone_number.Substring(found + 1);
            phone_number = phone_number.Insert(1, "(");
            phone_number = phone_number.Insert(5, ")");
            phone_number = phone_number.Insert(9, "-");
            phone_number = phone_number.Insert(12, "-");

            Console.Write($"Всё верно?: {surname} {name} {phonetic} {home_adress} {phone_number} (Y/N): ");

            answer = true;
            while (answer)
            {
                answertocheck = Console.ReadLine();


                if (answertocheck.EndsWith("Y") || answertocheck.EndsWith("N"))
                {
                    answer = false;
                }
                else { Console.Write("Не могу распознать ответ, введите его заново: "); }

            }

            found = answertocheck.IndexOf(":");
            answertocheck = answertocheck.Substring(found + 1);
            if (answertocheck == "Y")
            {
                AddNewClient(surname, name, phonetic, home_adress, phone_number);
            }
            else
            {
                Console.Clear();
            }
        }
        public int CheckClient(string surname)
        {
            WriteList();
            int found = 0;
            bool answer = true;
            int num = 0;
            string info = "";
            int k = 0;
            int[] index = new int[ClientsBase.GetLength(0)];
            for (int i = 0; i < ClientsBase.GetLength(0); i++)
            {
                if (surname == ClientsBase[i, 0])
                {
                    num++;
                    index[k] = i;

                    k++;
                }
            }
            Console.WriteLine($"Найдено совпадений : {num}");
            if (num != 0)
            {
                for (int i = 0; i < num; i++)
                {
                    for (int j = 0; j < ClientsBase.GetLength(1); j++)
                    {
                        info += $"  {ClientsBase[index[i], j]}  ";

                    }
                    Console.WriteLine($"\n {i + 1} - " + info);
                    info = "";
                }
                answer = true;
                while (answer)
                {


                    Console.Write($"Введите цифру, где указаны Вы(если вас нету в списке, нажмиите цифру {num + 1}): ");
                    string numerial = Console.ReadLine();
                    found = numerial.IndexOf(":");
                    numerial = numerial.Substring(found + 1);

                    if (Int32.TryParse(numerial, out num))
                    {
                        if (num == 3)
                            TakeDataClient(surname);
                        answer = false;

                        return index[num - 1];



                    }
                    else { Console.Write("Не могу распознать ответ, введите его заново: "); }
                }
            }
            else
            {

                Console.Write("Человека с такой фамилией нету у нас в базе, не желаете ли создать свою учетную запись?(Y/N): ");
                while (answer)
                {
                    string answertocreate = Console.ReadLine();
                    answertocreate = ti.ToUpper(answertocreate);



                    if (answertocreate.EndsWith("Y"))
                    {
                        TakeDataClient(surname);
                        answer = false;

                    }
                    else
                    {
                        if (answertocreate.EndsWith("N"))
                        {
                            Console.Clear();
                            Console.WriteLine("Тогда до свидания!");

                        }
                        else
                        {
                            Console.Write("Не могу распознать ответ, введите его заново: ");
                        }
                    }


                }

                return baseClient.Count();

            }
            return baseClient.Count();
        }
        public string TakeClientSurname()
        {
            Console.Clear();
            Console.Write("Введите вашу фамилию: ");
            string surnameToCheck = Console.ReadLine();
            surnameToCheck = ti.ToLower(surnameToCheck);
            surnameToCheck = ti.ToTitleCase(surnameToCheck);
            return surnameToCheck;
        }

        public void LK(int id)
        {
            Order orda = new Order();
            int index = 0;
            Console.Clear();
            Console.WriteLine("1- Изменить личные данные\n" +
                "2-Мои заказы\n" +
                "3-Новый заказ\n");

            string choise = Console.ReadLine();
            int found = choise.IndexOf(":");
            choise = choise.Substring(found + 1);

            bool answer = true;
            while (answer)
            {
                if (Int32.TryParse(choise, out index))
                {
                    if (index > 3)
                    {
                        Console.WriteLine("Неверный формат! ");
                    }
                    else
                    {
                        if (index < 1)
                        {
                            Console.WriteLine("Неверный формат! ");
                        }
                        else
                        {
                            answer = false;

                        }
                    }
                }
            }
            switch (index)
            {
                case 1:
                    ChangeUserDate(id);
                    break;
                case 2:
                    MyOrders(id);
                    break;
                case 3:
                    string surname =ClientBase[id, 0];
                    orda.AddNewOrder(surname, id);
                    break;
            }
        }
        public virtual void ChangeUserDate(int id)
        {
            Order ord = new Order();
            WriteList();
            int index = 0;
            Console.WriteLine("Что именно вы хотите изменить:\n" +
                "1-Фамилия\n" +
                "2-Имя\n" +
                "3-Отчество\n" +
                "4-Адрес\n" +
                "5-Номер телефона\n");
            Console.Write("\nВаш выбор : ");
            string choise = Console.ReadLine();
            int found = choise.IndexOf(":");
            choise = choise.Substring(found + 1);

            bool answer = true;
            while (answer)
            {
                if (Int32.TryParse(choise, out index))
                {
                    if (index > 5)
                    {
                        Console.WriteLine("Неверный формат! ");
                    }
                    else
                    {
                        if (index < 1)
                        {
                            Console.WriteLine("Неверный формат! ");
                        }
                        else
                        {
                            answer = false;

                        }
                    }
                }
            }
            switch (index)
            {
                case 1:
                    ChangeSurname(id);
                    
                    break;
                case 2:
                    ChangeName(id); 
                    break;
                case 3:
                    ChangeOtch(id);
                    break;
                case 4:
                    ChangeAdress(id);
                    break;
                case 5:
                    ChangePhone(id);
                    break;
            }
            MassToExcelREENTER();
            ord.ReEnterDateOrder(id);
        }
        public void ChangeName(int id)
        {
            Console.Clear();
            Console.Write("\nХорошо, введите новое имя: ");
            string newname = Console.ReadLine();
            int found = newname.IndexOf(":");
            newname = newname.Substring(found + 1);
            newname = ti.ToLower(newname);
            newname = ti.ToTitleCase(newname);
            
            ClientsBase[id, 1]= newname;
            
        }
        public void ChangeSurname(int id)
        {
            Console.Clear();
            Console.Write("Новая фамилия: ");


            string newsurname = Console.ReadLine();
            int found = newsurname.IndexOf(":");
            newsurname = newsurname.Substring(found + 1);
            newsurname = ti.ToLower(newsurname);
            newsurname = ti.ToTitleCase(newsurname);
            ClientsBase[id, 0] = newsurname;
        }
        public void ChangeOtch(int id)
        {
            Console.Clear();
            Console.Write("Новая фамилия: ");


            string newotch = Console.ReadLine();
            int found = newotch.IndexOf(":");
            newotch = newotch.Substring(found + 1);
            newotch = ti.ToLower(newotch);
            newotch = ti.ToTitleCase(newotch);
            ClientsBase[id, 2] = newotch;
        }
        public void ChangeAdress(int id)
        {
            char space = ' ';
            char newspace = '_';
            Console.Write("\nВведите ваш адрес: ");
            string newhome_adress = Console.ReadLine();
            int found = newhome_adress.IndexOf(":");
            newhome_adress = newhome_adress.Substring(found + 1);
            newhome_adress = newhome_adress.Replace(space, newspace);
            ClientsBase[id, 3] = newhome_adress;
        }
        public void ChangePhone(int id)
        {
            Console.Write("\nВведите ваш номер телефона: ");
            string newphone = Console.ReadLine();
            int found = newphone.IndexOf(":");
            newphone = newphone.Substring(found + 1);
            newphone = newphone.Insert(1, "(");
            newphone = newphone.Insert(5, ")");
            newphone = newphone.Insert(9, "-");
            newphone = newphone.Insert(12, "-");
            ClientsBase[id, 4] = newphone;
        }
        public void MassToExcelREENTER()
        {
            
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            var excelsheets = excelAppworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            Excel.Range xlRange = excelworksheet.UsedRange;
            for(int i =0; i < ClientsBase.GetLength(0); i++)
            {
                for(int j =0; j< ClientsBase.GetLength(1); j++)
                {
                    excelworksheet.Cells[i+2,j+1]= ClientsBase[i, j];
                    
                }
                /*for (char j = 'A'; j <= 'E'; j++)
                {
                    var excelcells = excelworksheet.get_Range($"{j}{(xlRange.Count / 5) + 1}");
                    
                    excelcells.Value2 = ClientsBase[i,k];
                    k++;
                }
                k = 0;*/
               
            }
            var excelappworkbooks = excelapp.Workbooks;
            var excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Save();
            excelapp.Quit();
            
            WriteList();
        }
        public void MyOrders(int id)
        {
            int k = 0;
            string text = "";
            Order or = new Order();
            string[,] OrdBase =or.GetOrderBase();
            int[] ind = new int[OrdBase.GetLength(0)];
            for (int i = 0; i< OrdBase.GetLength(0);i++)
            {
                if( id == Convert.ToInt32(OrdBase[i, 1]))
                {
                    ind[k] = i;
                    k++;
                }
            }
            Console.WriteLine($"На ваше имя найдено {k} заказов :");
            for (int i = 0; i < k; i++)
            {
                for (int j = 0; j < OrdBase.GetLength(1); j++)
                {
                    text += OrdBase[ind[i], j] + " ";
                }
                Console.WriteLine($"\n {i} - {text}\n");
                text = "";
            }
            Console.ReadKey();
        }
    }
}
