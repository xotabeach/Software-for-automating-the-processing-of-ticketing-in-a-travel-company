using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
namespace KursovayaRabota
{
    class Order : Clients, IRoute
    {
        int found = 0;
        string answertocreate = "";
        bool rez = true;
        bool answer = true;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        int index = 0;

        Route rot = new Route();

        public void AddNewOrder(string Surname, int idclient)
        {
            bool answer = true;
            string answertocreate = "";
            
            //Console.WriteLine("ID =" + idclient);
            Console.WriteLine();
            int indexMarsh = ChPref();
            Console.Write("\n Желаете ли продолжить оформление(Y/N): ");

            while (answer)
            {
                answertocreate = Console.ReadLine();
                answertocreate = answertocreate.ToUpper();

                if (answertocreate.EndsWith("Y") || answertocreate.EndsWith("N"))
                {
                    answer = false;
                }
                else { Console.Write("Не могу распознать ответ, введите его заново: "); }

            }
            int found = answertocreate.IndexOf(":");
            answertocreate = answertocreate.Substring(found + 1);
            if (answertocreate == "Y")
            {
                string datetoGo = Date();
                bool dopusl = DopUslugi();
                WriteOrder(idclient, indexMarsh, dopusl, datetoGo);
            }
            else
            {
                Console.Clear();
                Console.WriteLine("Вернёмся в начальное меню");
                System.Threading.Thread.Sleep(1000);
                Console.Clear();
            }


        }




        public int ChPref()
        {

            bool rez = true;

            int number;
            int index = 0;

            Route rot = new Route();

            int marshlength = rot.Marshbase;
            Console.WriteLine(marshlength);

            Console.WriteLine("Выберите по какому условию, вы хотите выбрать путёвку. \n " +
                    "Или же вы можете воспользоваться нашей рекомендацией(Для этого нажмите цифру 5): ");
            while (rez)
            {


                Console.WriteLine("1- Ввести страну, которую бы вы хотели посетить \n" +
                    "2- Ввести Продолжительность путевки, какая была бы для вас удобнее всего \n" +
                    "3- Ввести примерную цену, за которую вы бы хотели увидеть путёвку\n" +
                    "4- Ввести город, который вам интересен");
                Console.Write("Ваш выбор: ");
                string choose = Console.ReadLine();
                int chnum = choose.IndexOf(":");
                choose = choose.Substring(chnum + 1);

                if (Int32.TryParse(choose, out number))
                {
                    switch (number)
                    {
                        case 1:
                            index = OrderByCountry();
                            rez = false;
                            break;
                        case 2:
                            index = OrderByLong();
                            rez = false;
                            break;

                        case 3:
                            index = OrderByPrice();
                            rez = false;
                            break;
                        case 4:
                            index = OrderByCity();
                            rez = false;
                            break;
                        case 5:
                            index = OrderByRecomdation();
                            rez = false;
                            break;
                    }
                }
            }
            Console.WriteLine("INDEX = " + index);
            return index;
        }
        public void WriteInfo(string[] MARSH)
        {
            for (int i = 0; i < MARSH.Length; i++)
            {
                switch (i)
                {
                    case 0:
                        Console.WriteLine("\n Страна: " + MARSH[0]);
                        break;
                    case 1:

                        break;
                    case 2:
                        Console.WriteLine("\n Климат: " + MARSH[2]);
                        break;
                    case 3:
                        Console.WriteLine("\n Отель: " + MARSH[3]);
                        break;
                    case 4:
                        Console.WriteLine($"\n Проодолжительность тура: {MARSH[4]} недель");
                        break;
                    case 5:
                        Console.WriteLine($"\n Цена: {MARSH[5]} рублей");
                        break;
                }
            }
        }
        public void WriteOrder(int idUser, int idMarsh, bool dopusl, string date)
        {
            string marsh = "";

            marsh = rot.Marshbase2[idMarsh, 0];
            string user = ClientsBase[idUser, 0] + "_" + ClientsBase[idUser, 1] + "_" + ClientsBase[idUser, 2];

            int totalCost = Convert.ToInt32(rot.Marshbase2[idMarsh, 6]);
            Random r = new Random();
            int numberoforder = 0;
            int priceofmarsh = Convert.ToInt32(rot.Marshbase2[idMarsh, 6]);
            int skidka = 0;
            if (priceofmarsh > 40000)
            {
                skidka = 5;
            }
            if (dopusl)
                totalCost += 3000;
            if (priceofmarsh > 40000)
            {
                skidka = 5;
                totalCost -= totalCost * skidka / 100;
            }
            string newOrder = user + " " + marsh + " " + date + " " + priceofmarsh + " " + dopusl + " " + skidka + " " + totalCost;

            string[] order = newOrder.Split(" ");
            foreach (var gg in order)
                Console.WriteLine(gg);

            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            var excelsheets = excelAppworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
            Excel.Range xlRange = excelworksheet.UsedRange;
            int i = 0;
            for (char j = 'A'; j <= 'I'; j++)
            {
                var excelcells = excelworksheet.get_Range($"{j}{(xlRange.Count / 9) + 1}");
                numberoforder = r.Next(2000, 5000);
                if (i == 0)
                {
                    excelcells.Value2 = numberoforder;
                }
                else
                {
                    if(i==1 )
                    {
                        excelcells.Value2 = idUser;
                    }
                    else
                    {
                        excelcells.Value2 = order[i - 2];
                    }
                    

                }
                i++;
            }
            var excelappworkbooks = excelapp.Workbooks;
            var excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Save();
            excelapp.Quit();
        }
        public bool DopUslugi()
        {
            string answer = "";
            bool answercycle = true;
            bool cost = false;
            Console.Write("Наша компания предостовляет трансфер от аэропорта до вашего отеля. \n" +
                "Желаете ли приобрести данную опцию(цена 3000 рублей) (Y/N) : ");

            int found = 0;
            while (answercycle)
            {
                answer = Console.ReadLine();
                answer = answer.ToUpper();

                if (answer.EndsWith("Y") || answer.EndsWith("N"))
                {
                    found = answer.IndexOf(":");

                    answer = answer.Substring(found + 1);
                    if (answer == "Y")
                    {
                        cost = true;
                        answercycle = false;

                    }
                    else
                    {
                        if (answer == "N")
                        {
                            cost = false;
                            answercycle = false;
                        }
                    }

                }
                else { Console.Write("Не могу распознать ответ, введите его заново: "); }

            }
            return cost;
        }
        public string Date()
        {
            bool check = true;
            int month = 0;
            int day = 0;
            int year = 0;
            string date = "";

            DateTime thisday = DateTime.Today;
            Console.WriteLine(thisday.ToString());
            Console.WriteLine("Введите желаемую дату отъезда(дата должна быть не ранее семи дней от сегодняшней даты) : ");
            int found = 0;
            DateTime newdate;
            while (check)
            {
                Console.Write("\n Месяц(число) : ");
                string monthhh = Console.ReadLine();
                found = monthhh.IndexOf(":");
                monthhh = monthhh.Substring(found + 1);
                if (!Int32.TryParse(monthhh, out month))
                {

                    Console.WriteLine("Неверный формат!");
                }
                else
                {
                    Console.Write("\n День(число) : ");
                    string dayyyy = Console.ReadLine();
                    found = dayyyy.IndexOf(":");
                    dayyyy = dayyyy.Substring(found + 1);
                    if (!Int32.TryParse(dayyyy, out day))
                    {
                        Console.WriteLine("Неверный формат!");
                    }
                    else
                    {
                        Console.Write("\n Год(число) : ");
                        string yearrr = Console.ReadLine();
                        found = yearrr.IndexOf(":");
                        yearrr = yearrr.Substring(found + 1);
                        if (!Int32.TryParse(yearrr, out year))
                        {
                            Console.WriteLine("Неверный формат!");
                        }
                        else
                        {
                            date = day + "." + month + "." + year;

                            if (DateTime.TryParse(date, out newdate))
                            {
                                if ((newdate - thisday).TotalDays >= 7)
                                    check = false;
                                else
                                {
                                    Console.WriteLine("Некорректная дата!");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Некорректная дата!");
                            }

                        }


                    }
                }


            }
            return date;

        }
        static bool IsRealDate(int day, int month)

        {
            if (DateTime.DaysInMonth(DateTime.Today.Year, month) < day) return false;
            else return true;
        }

        public int OrderByLong()
        {
            int found = 0;




            int index = 0;

            Route rot = new Route();

            int marshlength = rot.Marshbase;
            int k = 0;
            int[] orderstolong = new int[rot.Marshbase];
            found = 0;
            bool cycle = true;
            Console.Write("Введите интересующую вас прододжительность путёвки(в неделях) : ");
            int howlongaweek = 0;
            while (cycle)
            {
                string element = Console.ReadLine();
                found = element.IndexOf(":");
                element = element.Substring(found + 1);
                if (!Int32.TryParse(element, out howlongaweek))
                {
                    Console.WriteLine("Неверный формат!");
                }
                else
                    cycle = false;
            }
            Console.WriteLine(howlongaweek);
            for (int i = 0; i < rot.Marshbase; i++)
            {
                if (howlongaweek == Convert.ToInt32(rot.Marshbase2[i, 5]))
                {

                    orderstolong[i] = i;
                    k++;
                }
            }
            if (k == 0)
            {
                Console.WriteLine("\n Извините, на данный момент у нашей тур. фирмы нету путёвок с продолжительностью {0} недель\n", howlongaweek);
            }

            string temptext2 = "";
            int[] newlongs = orderstolong.Except(new int[] { 0 }).ToArray();

            for (int i = 0; i < newlongs.Length; i++)
            {
                for (int j = 0; j < rot.Marshbase2.GetLength(1); j++)
                    temptext2 += rot.Marshbase2[Convert.ToInt32(newlongs[i]), j] + " ";
                temptext2 = i + "- " + temptext2;
                Console.WriteLine(temptext2);
                temptext2 = "";
            }
            Console.WriteLine("Выберите интересующую вас путёвку: ");
            int longg = 0;
            cycle = true;
            while (cycle)
            {
                string element = Console.ReadLine();
                found = element.IndexOf(":");
                element = element.Substring(found + 1);
                if (!Int32.TryParse(element, out longg))
                {
                    Console.WriteLine("Неверный формат!");
                }
                else
                    cycle = false;
            }
            index = Convert.ToInt32(newlongs[longg]);
            string texxt = rot.InfoAboutMarshbyIndex(index);
            string[] info = texxt.Split(" ");
            WriteInfo(info);
            return index;
        }
        public int OrderByCity()
        {




            int index = 0;

            Route rot = new Route();

            int marshlength = rot.Marshbase;
            Console.Write("Введите город, который вам хочется посетить: ");
            string findcities = "";
            string citywhatuserneed = Console.ReadLine();
            int temp = citywhatuserneed.IndexOf(":");
            citywhatuserneed = citywhatuserneed.Substring(temp + 1);
            for (int i = 0; i < marshlength; i++)
            {
                if (citywhatuserneed == Convert.ToString(rot.Marshbase2[i, 2]))
                {
                    findcities += i + " ";
                }
                if (i == marshlength - 1 && citywhatuserneed != rot.Marshbase2[i, 2] && findcities == "")
                {
                    Console.WriteLine($"\n Извините, но на данный момент в нашей фирме нету путёвок в город {citywhatuserneed}");
                    Console.WriteLine();
                }
            }

            findcities = findcities.Trim();
            string temptext = "";
            string[] cities = findcities.Split(" ");
            Console.WriteLine($"По вашему запросу найдено {cities.Length} путёвок: ");
            for (int i = 0; i < cities.Length; i++)
            {
                for (int j = 0; j < rot.Marshbase2.GetLength(1); j++)
                    temptext += rot.Marshbase2[Convert.ToInt32(cities[i]), j] + " ";
                temptext = i + "- " + temptext;
                Console.WriteLine(temptext);
                temptext = "";
            }
            Console.Write("Выберите номер понравившейся путёвки: ");
            string choose4 = Console.ReadLine();
            int chnum4 = choose4.IndexOf(":");
            choose4 = choose4.Substring(chnum4 + 1);
            index = Convert.ToInt32(cities[Convert.ToInt32(choose4)]);
            string text4 = rot.InfoAboutMarshbyIndex(index);
            string[] infobaoutmarsh4 = text4.Split(" ");
            WriteInfo(infobaoutmarsh4);
            return index;
        }
        public int OrderByRecomdation()
        {
            int found = 0;
            string answertocreate = "";

            bool answer = true;

            int index = 0;

            Route rot = new Route();

            int marshlength = rot.Marshbase;
            Random r = new Random();
            index = r.Next(1, marshlength);
            string text5 = rot.InfoAboutMarshbyIndex(index);
            string[] infobaoutmarsh5 = text5.Split(" ");

            Console.WriteLine($"\n Для вас мы рекомендуем путёвку в город {infobaoutmarsh5[1]} \n");
            WriteInfo(infobaoutmarsh5);
            Console.Write("\nПодходит ли вам наша рекомендация?(Y/N): ");
            while (answer)
            {
                answertocreate = Console.ReadLine();
                answertocreate = answertocreate.ToUpper();

                if (answertocreate.EndsWith("Y") || answertocreate.EndsWith("N"))
                {
                    answer = false;
                }
                else { Console.Write("Не могу распознать ответ, введите его заново: "); }

            }
            found = answertocreate.IndexOf(":");
            answertocreate = answertocreate.Substring(found + 1);
            if (answertocreate == "Y")
            {
                return index;
            }
            else
            {
                Console.Clear();
                Console.WriteLine("Тогда подберите тур по вашим желаниям");
            }
            return index;

        }
        public int OrderByPrice()
        {

            string answertocreate = "";

            bool answer = true;

            int index = 0;

            Route rot = new Route();

            int marshlength = rot.Marshbase;
            Console.Write("Введите цену, которая будет вам удобна: ");
            int findprices = 0;
            int[] prices = new int[rot.Marshbase];
            string pricewhatuserneed = Console.ReadLine();
            int temp3 = pricewhatuserneed.IndexOf(":");
            pricewhatuserneed = pricewhatuserneed.Substring(temp3 + 1);
            if (Int32.TryParse(pricewhatuserneed, out findprices))
            {
                for (int i = 0; i < rot.Marshbase; i++)
                {

                    if (Convert.ToInt32(rot.Marshbase2[i, 6]) < findprices || Convert.ToInt32(rot.Marshbase2[i, 6]) - findprices <= 4000)
                        prices[i] = Convert.ToInt32(rot.Marshbase2[i, 6]);
                }
            }
            Array.Sort(prices);
            Array.Reverse(prices);

            int[] newprices = prices.Except(new int[] { 0 }).ToArray();
            for (int i = 0; i < newprices.Length; i++)
                Console.WriteLine(newprices[i]);
            Console.WriteLine("Ближайшая стоимость путёвки составляет " + newprices[0] + " рублей. Показать полную информацию о ней?(Y/N): ");
            while (answer)
            {
                answertocreate = Console.ReadLine();
                answertocreate = answertocreate.ToUpper();

                if (answertocreate.EndsWith("Y") || answertocreate.EndsWith("N"))
                {
                    answer = false;
                }
                else { Console.Write("Не могу распознать ответ, введите его заново: "); }

            }

            int found3 = answertocreate.IndexOf(":");
            string text3 = "";
            answertocreate = answertocreate.Substring(found3 + 1);
            if (answertocreate == "Y")
            {
                for (int i = 0; i < marshlength; i++)
                {
                    if (newprices[0] == Convert.ToInt32(rot.Marshbase2[i, 6]))
                    {
                        index = i;
                    }
                }
                text3 = rot.InfoAboutMarshbyIndex(index);
            }
            else
            {
                Console.Clear();
            }
            string[] marsh3 = text3.Split();
            WriteInfo(marsh3);
            return index;
        }
        public int OrderByCountry()
        {
            int index = 0;
            string text = "";
            string choise = "";
            bool answer = true;
            int choiseint = 0;
            int k = 0;
            Route rot = new Route();
            TextInfo ti = CultureInfo.CurrentCulture.TextInfo;
            int marshlength = rot.Marshbase;
            Console.Write("Введите страну, которую вам хочется посетить: ");
            string country = "";
            int[] countries = new int[rot.Marshbase];
            string countrywhatuserneed = Console.ReadLine();
            int temp = countrywhatuserneed.IndexOf(":");
            country = countrywhatuserneed.Substring(temp + 1);
            country = country.ToLower();

            country = ti.ToTitleCase(country);
            for (int i = 0; i < rot.Marshbase; i++)
            {
                if (country == rot.Marshbase2[i, 1])
                {
                    countries[k] = i;
                    k++;
                }
            }
            int[] newcountries = countries.Except(new int[] { 0 }).ToArray();
            Console.WriteLine($"\nПо вашему запросу найдено {k} марщрутов в {country} : \n");
            for (int i = 0; i < newcountries.Length; i++)
            {
                for (int j = 0; j < rot.Marshbase2.GetLength(1); j++)
                {
                    text += rot.Marshbase2[newcountries[i], j] + " ";
                }
                Console.WriteLine($"\n {i} - {text}\n");
                text = "";
            }
            Console.WriteLine("\n Выберите понравившийся вариант : ");
            choise = Console.ReadLine();
            found = choise.IndexOf(":");
            choise = choise.Substring(found + 1);
            string info = "";
            while (answer)
            {
                if (Int32.TryParse(choise, out choiseint))
                {
                    if (choiseint > newcountries.Length)
                    {
                        Console.WriteLine("Неверный формат! ");
                    }
                    else
                    {
                        if (choiseint < 0)
                        {
                            Console.WriteLine("Неверный формат! ");
                        }
                        else
                        {
                            index = newcountries[choiseint];
                            info = rot.InfoAboutMarshbyIndex(newcountries[choiseint]);
                            
                            answer = false;
                        }
                    }
                }
            }
            string[] infotext = info.Split(" ");
            WriteInfo(infotext);
            return index;
        }
        public void CheckOrder()
        {
            
            string name = "";
            string surname;
            string otch = "";
            string sur_name_order = "";
            surname = TakeClientSurname();
            int id = CheckClient(surname);
            int index = 0;
            name = ClientsBase[id, 1];
            otch = ClientsBase[id, 2];
            sur_name_order = surname + "_" + name + "_" + otch;
            string[,] Base = GetOrderBase();
            for (int i = 0; i < Base.GetLength(0); i++)
            {
                if (sur_name_order == Base[i, 2] || id == Convert.ToInt32(Base[i,1]))
                {
                    index = i;
                    break;
                }
            }
            WriteInfoBaotOrder(Base, index);

            if (Convert.ToDateTime(Base[index, 4]) > DateTime.Now)
            {
                Console.WriteLine(" \nСтатус :Оформляется");
            }
            else
            {
                Console.WriteLine(" \nСтатус :Исполнена");
            }
            Console.ReadKey();
        }
        public string[,] GetOrderBase()
        {
            string text = "";
            int h = 0;
            int k = 0;
            string[,] OrderBase;
            List<string> orderBase = new List<string>();
            //StreamReader read_cl_base = new StreamReader("ClientsBase.txt");
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            excelsheets = excelAppworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
            Excel.Range xlRange = excelworksheet.UsedRange;
            OrderBase = new string[xlRange.Count / 9, 9];
            for (int i = 2; i <= xlRange.Count / 9; i++)
            {
                for (char j = 'A'; j <= 'I'; j++)
                {

                    //Console.WriteLine(xlRange.Count);
                    var excelcells = excelworksheet.get_Range($"{j}{i}", Type.Missing);
                    string sStr = Convert.ToString(excelcells.Value2);

                    text += sStr + " ";
                    OrderBase[h, k] = sStr;
                    k++;

                }
                k = 0;
                h++;
                orderBase.Add(text);
                text = "";

            }
            excelapp.Quit();
            return OrderBase;
        }
        public void WriteInfoBaotOrder(string[,] Base, int index)
        {
            for (int i = 0; i < Base.GetLength(1); i++)
            {
                switch (i)
                {
                    case 0:
                        Console.WriteLine($" Номер заказа: {Base[index, 0]}\n");
                        break;
                    case 1:
                        Console.WriteLine($" ФИО:  {Base[index, 2]}\n");
                        break;
                    case 2:
                        Console.WriteLine($" Номер маршрута:  {Base[index, 3]}\n");
                        break;
                    case 3:
                        Console.WriteLine($" Дата отправления:  {Base[index, 4]}\n");
                        break;
                    case 4:
                        Console.WriteLine($" Цена путевки:  {Base[index, 5]}\n");
                        break;
                    case 5:
                        Console.WriteLine($" Наличие доп. услуги:  {Base[index, 6]}\n");
                        break;
                    case 6:
                        Console.WriteLine($" Скидка:  {Base[index, 7]}%\n");
                        break;
                    case 7:
                        Console.WriteLine($" Итоговая стоимость:  {Base[index, 8]}\n");
                        break;

                }
            }
        }
        public bool CheckToChange(int i)
        {
            bool check = false;
            string[,] ordBase = GetOrderBase();
            DateTime d2 = Convert.ToDateTime(ordBase[i, 4]);
            
            if ((d2 - DateTime.Now ).TotalDays >= 4)
            {
                check = true;
                Console.WriteLine("\n Изменение возможно!\n");
            }
            else
            {
                Console.WriteLine("\n Изменение невозможно!");
            }
            return check;
        }
        public int TakeNumberofChange(bool check)
        {
            string choise;
            int index = 0;
            if (check){
                Console.Clear();
                Console.WriteLine("Выберите из списка, что именно вы хотите изменить в вашем заказе :");
                Console.WriteLine("\n1- Маршрут\n" +
                    "2- Дата отьезда\n" +
                    "3- Доп.Услуга\n");
                Console.Write("Ваш выбор: ");
                choise = Console.ReadLine();
                found = choise.IndexOf(":");
                choise = choise.Substring(found + 1);

                answer = true;
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
                return index;
            }
            else
            {
                //Console.Clear();
            }
            return index;


        }
        public void ChangeOrderData(int changeindex, int id)
        {
            string answ = "";
            bool answercycle = true;
            string choise = "";
            int found = 0;
            int choiseint = 0;
            IRoute newmarsh = new Route();
            Route mrsh = new Route();
            bool cycle = true;
            bool dop = false;
            while (cycle)
            {
                Console.Clear();
                switch (changeindex)
                {
                    
                    case 1:
                        Console.WriteLine("\nВыберите новый маршрут \n");
                        
                        newmarsh.PrintInfo();
                        Console.Write("\nВаш выбор : ");
                        choise = Console.ReadLine();
                        found = choise.IndexOf(":");
                        choise = choise.Substring(found + 1);
                        
                        if (Int32.TryParse(choise, out choiseint))
                            {
                                if( choiseint> mrsh.Marshbase)
                                {
                                    Console.WriteLine("\nнекорректный формат!\n");
                                }
                                else
                                {
                                string nm = mrsh.Marshbase2[choiseint, 0];
                                    OrderToExcel(id,nm, changeindex);
                                    
                                }
                            }
                        
                        
                        break;
                    case 2:
                        Console.WriteLine("\nВыберите новую дату отъезда:\n");

                        
                        string text = GetNewDate();
                        OrderToExcel(id, text, changeindex);
                        cycle = false;

                        break;
                    case 3:
                        Console.Write("Подключить доп. услугу?(Y/N): ");
                        string[,] ordbase = GetOrderBase(); 
                        while (answercycle)
                        {
                            answ = Console.ReadLine();
                            answ = answ.ToUpper();

                            if (answ.EndsWith("Y") || answ.EndsWith("N"))
                            {
                                found = answ.IndexOf(":");

                                answ = answ.Substring(found + 1);
                                if (answ == "Y")
                                {
                                    dop = true;
                                    OrderToExcel(id, Convert.ToString(dop), changeindex);
                                    answercycle = false;

                                }
                                else
                                {
                                    if (answ == "N")
                                    {
                                        dop = false;
                                        OrderToExcel(id, Convert.ToString(dop), changeindex);
                                        string newcost =ordbase[id, 5];
                                        OrderToExcel(id, newcost, changeindex+1);
                                        answercycle = false;
                                    }
                                }

                            }
                            else { Console.Write("Не могу распознать ответ, введите его заново: "); }

                        }
                        break;
                }
                cycle = false;
            }
        }
        public int TakeNumberofOrder()
        {
            Console.WriteLine("\nИзменить данные в вашем заказе возможно только при условии, что разница между датой отъезда и датой изменения меньше 4 суток!\n");
            int ind = 0;
            string[,] orBase= GetOrderBase();
            int number = 0;
            Console.Write(" Введите ваш номер заказа: ");
            string choise = Console.ReadLine();
            int found = choise.IndexOf(":");
            choise = choise.Substring(found + 1);
            
            

            for (int i = 0; i < orBase.GetLength(0); i++)
            {
                 if (Int32.TryParse(choise, out number))
                {
                    if (number == Convert.ToInt32(orBase[i, 0]))
                    {
                        ind = i;
                        break;
                    }
                }
            }
            return ind;
        }
        public void OrderToExcel(int index, string value, int changeindex)
        {
            
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            var excelsheets = excelAppworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
            Excel.Range xlRange = excelworksheet.UsedRange;
            
            switch (changeindex)
            {
                case 1:
                    excelworksheet.Cells[index + 2, changeindex + 3] = value;
                    /*var excelcells = excelworksheet.get_Range($"{index+1}{changeindex+1}");
                    excelcells.Value2 = value;*/

                    var excelappworkbooks = excelapp.Workbooks;
                    var excelappworkbook = excelappworkbooks[1];
                    excelappworkbook.Save();
                    excelapp.Quit();
                    break;
                case 2:
                    excelworksheet.Cells[index + 2, changeindex + 4] = value;
                    /*excelcells = excelworksheet.get_Range($"{index}{changeindex+2}");
                    excelcells.Value2 = value;*/

                    excelappworkbooks = excelapp.Workbooks;
                    excelappworkbook = excelappworkbooks[1];
                    excelappworkbook.Save();
                    excelapp.Quit();


                    break;
                case 3:
                    /* excelcells = excelworksheet.get_Range($"{index+1}{changeindex+3}");*/
                    excelworksheet.Cells[index + 2, changeindex + 4] = value;
                    excelappworkbooks = excelapp.Workbooks;
                    excelappworkbook = excelappworkbooks[1];
                    excelappworkbook.Save();
                    excelapp.Quit(); 
                    ////excelcells.Value2 = ;
                    break;
                case 4:
                    excelworksheet.Cells[index + 2, changeindex + 5]= value;
                    excelappworkbooks = excelapp.Workbooks;
                    excelappworkbook = excelappworkbooks[1];
                    excelappworkbook.Save();
                    excelapp.Quit();
                    break;
            }
            /*var excelappworkbooks = excelapp.Workbooks;
            var excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Save();
            excelapp.Quit();*/


            }
            public string GetNewDate()
        {
            bool check = true;
            int month = 0;
            int day = 0;
            int year = 0;
            string date = "";

            DateTime thisday = DateTime.Today;
            int found = 0;
            DateTime newdate;
            while (check)
            {
                Console.Write("\n Месяц(число) : ");
                string monthhh = Console.ReadLine();
                found = monthhh.IndexOf(":");
                monthhh = monthhh.Substring(found + 1);
                if (!Int32.TryParse(monthhh, out month))
                {

                    Console.WriteLine("Неверный формат!");
                }
                else
                {
                    Console.Write("\n День(число) : ");
                    string dayyyy = Console.ReadLine();
                    found = dayyyy.IndexOf(":");
                    dayyyy = dayyyy.Substring(found + 1);
                    if (!Int32.TryParse(dayyyy, out day))
                    {
                        Console.WriteLine("Неверный формат!");
                    }
                    else
                    {
                        Console.Write("\n Год(число) : ");
                        string yearrr = Console.ReadLine();
                        found = yearrr.IndexOf(":");
                        yearrr = yearrr.Substring(found + 1);
                        if (!Int32.TryParse(yearrr, out year))
                        {
                            Console.WriteLine("Неверный формат!");
                        }
                        else
                        {
                            date = day + "." + month + "." + year;

                            if (DateTime.TryParse(date, out newdate))
                            {
                                check = false;
                                return date;
                                
                                
                            }
                            else
                            {
                                Console.WriteLine("Некорректная дата!");
                            }

                        }


                    }
                }


            }
            return date;
        }
        public void ReEnterDateOrder(int id)
        {
            WriteList();
            bool check = false;
            int index = 0;
            string[,] BaseMass = GetOrderBase();
            for(int i =0; i< BaseMass.GetLength(0); i++)
            {
                if( id == Convert.ToInt32(BaseMass[i, 1]))
                {
                    index = i;
                    check = true;
                }
            }
            if(check)
            {
                string newdate = ClientsBase[id, 0] + "_" + ClientsBase[id, 1] + "_" + ClientsBase[id, 2];
                BaseMass[index, 2] = newdate;
                Console.WriteLine(newdate);
                BaseToExcell(BaseMass);
            }
            
        }
        public void BaseToExcell(string[,] baseOrder)
        {
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            var excelsheets = excelAppworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
            Excel.Range xlRange = excelworksheet.UsedRange;
           
            for(int i =0; i< baseOrder.GetLength(0); i++)
            {
                for(int j = 0; j < baseOrder.GetLength(1); j++)
                {
                excelworksheet.Cells[i + 2, j + 1] = baseOrder[i, j];
                    
                    
                }
                
            }
            
            var excelappworkbooks = excelapp.Workbooks;
            var excelappworkbook = excelappworkbooks[1];
            excelappworkbook.Save();
            excelapp.Quit();
        }
    }
}