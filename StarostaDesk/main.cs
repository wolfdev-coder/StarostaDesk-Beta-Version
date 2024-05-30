using OfficeOpenXml;
using System.Data;
using System.Data.SQLite;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.InputFiles;
using Telegram.Bot.Types.ReplyMarkups;
namespace main
{
    class mainBot
    {
        #region Переменные паблики и приваты
        public static readonly SQLiteConnection DB = new SQLiteConnection(database.connection);
        private static TelegramBotClient Bot;
        public static int permission = 0;
        public static string curatorId { get; set; }
        public static int ReplyId = 0;
        public static string fileSource = @"D:\MyFiles\Отчет.xlsx";
        public static string curatorName { get; set; }
        public static string curatorCours { get; set; }
        public static List<string> studentNames = new List<string>();
        public static string starostaName { get; set; }
        public static string studentName { get; set; }

        public static string studentDay { get; set; }
        public static string studentHours { get; set; }
        public static string starostaCours { get; set; }
        public static string monthTable { get; set; }
        public static string starostaId { get; set; }
        //public string[] StudentsList { get; private set; }
        #endregion
        public static void Main(string[] args)
        {

            DB.OpenAsync();
            Bot = new TelegramBotClient($"{programm.token}");
            Bot.OnMessage += Bot_OnMessageReceived;
            Bot.StartReceiving();
            Console.WriteLine("Bot started");
            Console.ReadKey();

        }

        [Obsolete]
        private static async void Bot_OnMessageReceived(object? sender, MessageEventArgs e)
        {
            #region Кнопачки
            try
            {
                
                var message = e.Message;
                Console.WriteLine(message.From.Id + " " + message.Text);
                
                var starostaBtn = new ReplyKeyboardMarkup
                {
                    Keyboard = new[]
    {
                    new []
                    {
                        new KeyboardButton("Пропуск🔍"),
                        new KeyboardButton("Команды❓"),
                        new KeyboardButton("Отчет🟢")
                    }
                },
                    ResizeKeyboard = true
                };
                var curatorBtn = new ReplyKeyboardMarkup
                {
                    Keyboard = new[]
{
                    new []
                    {
                        new KeyboardButton("Пропуск🔍"),
                        new KeyboardButton("Команды❓"),
                        new KeyboardButton("Отчет🟢")
                    },
                    new []
                    {
                        new KeyboardButton("Создать месяц"),
                        new KeyboardButton("Создать группу"),
                        new KeyboardButton("Перенести")
                    }

                },
                    ResizeKeyboard = true
                };
                var starostaHoursBtn = new ReplyKeyboardMarkup
                {
                    Keyboard = new[]
    {
                    new []
                    {
                        new KeyboardButton("2"),
                        new KeyboardButton("4"),
                        new KeyboardButton("6"),
                        new KeyboardButton("8")
                    }
                },
                    ResizeKeyboard = true
                };

                #endregion
                switch (message.Text)
                {
                    case "/start":
                        await Bot.SendTextMessageAsync(message.From.Id, "Здравствуйте!", replyMarkup: starostaBtn);
                        break;
                    case "Команды❓":
                        await Bot.SendTextMessageAsync(message.From.Id, "Здравствуйте! Вот мой список команд:\n\nКураторам:\n1)Создать группу - создает таблицу группы, привязанной к вашему профилю" +
                            "\n2)Создать месяц - создает таблицу для отчета посещаемости студентов\n3)Перенести - переносит с таблицы с ФИО студеннтами в таблицу с отчетом посещаемости" +
                            "\n4)Староста - добавляет старосту в таблицу группы" +
                            "\n\nСтаростам:\n1)Пропуск - вносит прогул за ТЕКУЩИЙ день\n2)Прогул - вносит прогул за любой указанный вами день\n3)Отчет - создает отчет за месяц и отправляет вам", replyMarkup: starostaBtn);
                        break;
                }
                #region Все команды
                if (message.Text.Contains("Пропуск🔍"))
                {
                    Together();
                    LoadStarosta(message.From.Id.ToString());
                    if (starostaId == message.From.Id.ToString())
                    {
                        studentNames.Clear();
                        SQLiteCommand cmd = new SQLiteCommand($"SELECT name FROM [{starostaCours}]", DB);
                        SQLiteDataReader reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            string studentName = reader.GetString(0);
                            if (studentName != null)
                            {
                                studentNames.Add(studentName);
                            }
                        }
                        reader.Close();
                        string table = $"{monthTable}{starostaCours}";
                        var uttons = studentNames.Select(name => new KeyboardButton[] { new KeyboardButton(name) }).ToArray();
                        var namesKb = new ReplyKeyboardMarkup(uttons);
                        await Bot.SendTextMessageAsync(message.From.Id, "🔍Выберите студента:", replyMarkup: namesKb);
                        permission = 1;

                    }
                }
                if (message.Text.StartsWith("Обновление") && message.From.Id == 1251534440)
                {
                    string target_id;
                    string[] parts = message.Text.Split('/');
                    string updatemessage = parts[1];
                    SQLiteCommand AllSelectCurator = new SQLiteCommand("SELECT * FROM curators", DB);
                    SQLiteDataReader readerCurator = AllSelectCurator.ExecuteReader();
                    while (readerCurator.Read())
                    {
                        target_id = readerCurator.GetString(2);
                        await Bot.SendTextMessageAsync(target_id, $"🔄Обновление!\n {updatemessage}");
                    }
                    Thread.Sleep(700);
                    await Bot.SendTextMessageAsync(1251534440, "Отправил кураторам, начинаю отправку старостам..");
                    SQLiteCommand AllSelectStarosta = new SQLiteCommand("SELECT * FROM starosta", DB);
                    SQLiteDataReader readerStarosta = AllSelectStarosta.ExecuteReader();
                    while(readerStarosta.Read())
                    {
                        target_id = readerStarosta.GetString(2);
                        await Bot.SendTextMessageAsync(target_id, $"🔄Обновление!: {updatemessage}");
                    }
                    await Bot.SendTextMessageAsync(1251534440, "Выслал всем)");

                }
                if (studentNames.Contains(message.Text) && permission == 1 && message.From.Id.ToString() == starostaId)
                {
                    studentName = message.Text;
                    await Bot.SendTextMessageAsync(message.From.Id, "❓Сколько прогулял данный студент?", replyMarkup: starostaHoursBtn);
                    permission = 2;

                }
                if (message.Text.Contains("2") && permission == 2 || message.Text.Contains("4") && permission == 2 || message.Text.Contains("6") && permission == 2 || message.Text.Contains("8") && permission == 2)
                {
                    permission = 0;
                    Together();
                    LoadStarosta(message.From.Id.ToString());
                    string table = $"{monthTable}{starostaCours}";
                    AddStudentNull(table, DB, message.Text, studentName);
                    await Bot.SendTextMessageAsync(message.From.Id, $"👨‍🎓Студент: {studentName} \n🕒Пропустил {message.Text} ч\n✅Прогул занесен в таблицу\n{starostaName} хорошая работа!\n\n❗Если студент прогулял еще пару, то выберите вариант больше, чем {message.Text}", replyMarkup: starostaBtn);
                }

                if (message.Text.Contains("Создать месяц"))
                {
                    string currentMonth = DateTime.Now.ToString("MMMM", new System.Globalization.CultureInfo("en-US"));
                    string monthTable = currentMonth.ToLower();

                    SQLiteCommand command1 = new SQLiteCommand("SELECT name, cours, tg_id FROM curators", DB);
                    SQLiteDataReader sqlite_datareader = command1.ExecuteReader();
                    while (sqlite_datareader.Read())
                    {
                        curatorName = sqlite_datareader.GetString(0);
                        curatorCours = sqlite_datareader.GetString(1);
                        string tg_id = sqlite_datareader.GetString(2);
                        Console.WriteLine(curatorName + " " + curatorCours);
                        if (tg_id == message.From.Id.ToString())
                        {
                            string tableName = $"{monthTable}{curatorCours}";
                            CreateTableNone(tableName, DB);
                            await Bot.SendTextMessageAsync(message.From.Id, "✅Таблица " + tableName + " создана!\n🔴Это таблица для контроля посещаемости студентов\n🔴Староста теперь работает с новым месяцем", replyMarkup: starostaBtn);
                        }
                    }
                    sqlite_datareader.Close();
                }
                if (message.Text.Contains("Запросить"))
                { 
                    await Bot.SendTextMessageAsync(message.From.Id, $"🕐Ждем ответа от администратора\n⚪Как только вас примут, я отправлю вам сообщение с вашей должностью");
                    await Bot.SendTextMessageAsync(1251534440, $"⚪Пользователь с айди {message.From.Id}, хочет группу");
                }
                if (message.Text.StartsWith("Куратор") && message.From.Id == 1251534440)
                {
                    string[] parts = message.Text.Split('/');
                    curatorId = parts[1];
                    curatorName = parts[2];
                    curatorCours = parts[3];
                    SQLiteCommand command = new SQLiteCommand("INSERT INTO curators (name, cours, tg_id) VALUES (@name, @cours, @tg_id)", DB);
                    command.Parameters.AddWithValue("@tg_id", curatorId);
                    command.Parameters.AddWithValue("@name", curatorName);
                    command.Parameters.AddWithValue("@cours", curatorCours);
                    command.ExecuteNonQuery();
                    await Bot.SendTextMessageAsync(message.From.Id, $"👼Куратор с айди {curatorId} \n🔴Имя: {curatorName} \n🔴Группа {curatorCours} \n✅Куратор успешно зарегистрирован!");
                    await Bot.SendTextMessageAsync(curatorId, $"✅Вы успешно зарегистрированы!\n🔴Ваше имя: {curatorName}\n🔴Ваша группа {curatorCours}\n👼Должность: Куратор");

                }
                if (message.Text.StartsWith("Староста"))
                {
                    
                    string[] parts = message.Text.Split('/');
                    starostaId = parts[1];
                    starostaName = parts[2];
                    starostaCours = parts[3];
                    SQLiteCommand command = new SQLiteCommand("INSERT INTO starosta (name, cours, tg_id) VALUES (@name, @cours, @tg_id)", DB);
                    command.Parameters.AddWithValue("@tg_id", starostaId);
                    command.Parameters.AddWithValue("@name", starostaName);
                    command.Parameters.AddWithValue("@cours", starostaCours);
                    command.ExecuteNonQuery();
                    await Bot.SendTextMessageAsync(message.From.Id, $"👨‍🎤Староста с айди {starostaId} \n🟢Имя: {starostaName} \n🟢Группа {starostaCours} \n✅Староста успешно зарегистрирован!");
                    await Bot.SendTextMessageAsync(starostaId, $"✅Вы успешно зарегистрированы!\n🟢Ваше имя: {starostaName}\n🟢Ваша группа {starostaCours}\n👨‍🎤Должность: Староста");

                }
                if (message.Text.StartsWith("Отправь."))
                {
                    string[] parts = message.Text.Split('.');
                    var id = int.Parse(parts[1]);
                    var replyMessage = parts[2];
                    await Console.Out.WriteLineAsync($"айди: {id} мессаге: {replyMessage}, администратор: {message.From.Id}");
                    await Bot.SendTextMessageAsync(id, $"💬Новое сообщение!\n\n💭{replyMessage}\n\n✅Чтобы ответить на данное сообщение, введите следующую команду: Отправь.{message.From.Id}.текст сообщения");
                    await Bot.DeleteMessageAsync(message.From.Id, message.MessageId);
                    await Bot.SendTextMessageAsync(message.From.Id, "💌Сообщение отправлено!");
                }
                if (message.Text.Contains("Создать группу"))
                {
                    LoadCurator(message.From.Id.ToString());
                    string tableName = curatorCours;
                    CreateTableGroup(tableName, DB);
                    await Bot.SendTextMessageAsync(message.From.Id, $"✅Группа: {curatorCours} успешно создана!\n🔴Теперь добавьте в нее студентов с помощью команды: Добавить/ФИО Студента");
                }
                if (message.Text.Contains("Отчет🟢"))
                {
                    Together();
                    SQLiteCommand command = new SQLiteCommand("SELECT name, cours, tg_id FROM starosta", DB);
                    SQLiteDataReader datareader = command.ExecuteReader();
                    while (datareader.Read())
                    {
                        string name = datareader.GetString(0);
                        string cours = datareader.GetString(1);
                        string tg_id = datareader.GetString(2);
                        if (message.From.Id.ToString() == tg_id)
                        {
                            string table = $"{monthTable}{cours}";
                            ExportToExcel(fileSource, table);
                            var fileStream = System.IO.File.OpenRead(fileSource);
                            InputOnlineFile inputOnlineFile = new InputOnlineFile(fileStream, "Отчет.xlsx");
                            await Bot.SendDocumentAsync(message.From.Id, inputOnlineFile);

                        }
                    }
                    datareader.Close();

                }
                if (message.Text.Contains("Добавить/"))
                {
                    string[] parts = message.Text.Split('/');
                    studentName = parts[1];
                    LoadCurator(message.From.Id.ToString());
                    await Bot.SendTextMessageAsync(message.From.Id, $"✅Студент: {studentName} - добавлен в таблицу группы {curatorCours}");
                    string table = curatorCours;
                    AddStudent(table, DB, studentName);
                }
                if (message.Text.StartsWith("Удалить"))
                {
                    string[] parts = message.Text.Split('/');
                    var deleteId = parts[1];
                    LoadCurator(message.From.Id.ToString());
                    if (message.From.Id.ToString() == curatorId) 
                    { 
                        SQLiteCommand cmd = new SQLiteCommand("DELETE FROM starosta WHERE tg_id=@tg_id", DB);
                        cmd.Parameters.AddWithValue("@tg_id", deleteId);
                        cmd.ExecuteNonQuery();
                        await Bot.SendTextMessageAsync(message.From.Id, "Староста удален");
                        await Bot.SendTextMessageAsync(deleteId, "Вы были отключены от системы, всего хорошего!");
                    }
                }
                if (message.Text.StartsWith("Прогул"))
                {

                    string[] parts = message.Text.Split('/');
                    studentName = parts[1];
                    studentDay = parts[2];
                    studentHours = parts[3];
                    Together();
                    LoadStarosta(message.From.Id.ToString());
                    string table = $"{monthTable}{starostaCours}";
                    AddStudentNullWithDay(table, DB, studentHours, studentName, studentDay);
                    await Bot.SendTextMessageAsync(message.From.Id, $"👨‍🎓Студент: {studentName} \n🕒Пропустил: {studentHours} часов\n🟢Число месяца:{studentDay}\n✅Успешно внесено в таблицу");
                }
                if (message.Text.Contains("Перенести"))
                {
                    Together();
                    LoadCurator(message.From.Id.ToString());
                    string table = $"{monthTable}{curatorCours}";
                    SQLiteCommand command1 = new SQLiteCommand($"SELECT name FROM [{curatorCours}]", DB);
                    SQLiteDataReader reader = command1.ExecuteReader();

                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            studentName = reader.GetString(0);
                            SQLiteCommand command2 = new SQLiteCommand($"INSERT INTO [{table}] (name) VALUES ('{studentName}')", DB);
                            command2.ExecuteNonQuery();
                            await Bot.SendTextMessageAsync(message.From.Id, $"✅Добавил: {studentName} в таблицу {table}");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                await Bot.SendTextMessageAsync(1251534440, "🛑У меня случилась следующая ошибка: \n" + ex.Message + "\n\nВ чате: " + e.Message.From.Id);
            }
        }
        
        #endregion
        
        #region Методы

        private static void CreateTableNone(string tableName, SQLiteConnection connection)
        {
            try
            {
                using (SQLiteCommand command = new SQLiteCommand(connection))
                {
                    command.CommandText = $"CREATE TABLE [{tableName}] (id INTEGER PRIMARY KEY, name TEXT, day1 INTEGER, day2 INTEGER, day3 INTEGER, day4 INTEGER, day5 INTEGER, day6 INTEGER, day7 INTEGER, day8 INTEGER, day9 INTEGER, day10 INTEGER, day11 INTEGER, day12 INTEGER, day13 INTEGER, day14 INTEGER, day15 INTEGER, day16 INTEGER, day17 INTEGER, day18 INTEGER, day19 INTEGER, day20 INTEGER, day21 INTEGER, day22 INTEGER, day23 INTEGER, day24 INTEGER, day25 INTEGER, day26 INTEGER, day27 INTEGER, day28 INTEGER, day29 INTEGER, day30 INTEGER, day31 INTEGER)";
                    command.ExecuteNonQuery();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void Together()
        {
            string currentMonth = DateTime.Now.ToString("MMMM", new System.Globalization.CultureInfo("en-US"));
            monthTable = currentMonth.ToLower();
        }
        static void ExportToExcel(string excelFile, string tableName)
        {
            try
            {
                SQLiteCommand command = new SQLiteCommand($"SELECT * FROM [{tableName}]", DB);
                SQLiteDataReader rdr = command.ExecuteReader();
                var dataTable = new DataTable();
                dataTable.Load(rdr);
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[1, j + 3].Value = (j + 1);
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                        }
                    }

                    package.SaveAs(new System.IO.FileInfo(excelFile));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static void CreateTableGroup(string tableName, SQLiteConnection connection)
        {
            using (SQLiteCommand command = new SQLiteCommand(connection))
            {
                command.CommandText = $"CREATE TABLE IF NOT EXISTS [{tableName}] (name TEXT)";
                command.ExecuteNonQuery();
            }
        }
        public static async void LoadStarosta(string tg_id)
        {
            try
            {
                SQLiteCommand command = new SQLiteCommand($"SELECT name, cours, tg_id FROM starosta WHERE tg_id='{tg_id}'", DB);
                SQLiteDataReader datareader = command.ExecuteReader();
                datareader.Read();
                starostaName = datareader.GetString(0);
                starostaCours = datareader.GetString(1);
                starostaId = datareader.GetString(2);
                await Console.Out.WriteLineAsync(starostaName + " \n" + starostaId + "\n" + starostaCours);
                datareader.Close(); 
            }
            catch (Exception ex)
            {
                await Console.Out.WriteLineAsync(ex.Message);
            }
        }
        public static async void LoadCurator(string tg_id)
        {
            try
            {
                SQLiteCommand command1 = new SQLiteCommand($"SELECT name, cours, tg_id FROM curators WHERE tg_id='{tg_id}'", DB);
                SQLiteDataReader sqlite_datareader = command1.ExecuteReader();
                sqlite_datareader.Read();
                curatorName = sqlite_datareader.GetString(0);
                curatorCours = sqlite_datareader.GetString(1);
                curatorId = sqlite_datareader.GetString(2);
                sqlite_datareader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static void AddStudent(string tableName, SQLiteConnection connection, string name)
        {
            using (SQLiteCommand command = new SQLiteCommand(connection))
            {
                command.CommandText = $"INSERT INTO [{tableName}] (name) VALUES (@name)";
                command.Parameters.AddWithValue("@name", name);
                command.ExecuteNonQuery();
            }
        }
        private static void AddStudentNull(string tableName, SQLiteConnection connection, string hours, string student)
        {
            int day = DateTime.Today.Day;
            string today = $"day{day}";
            using (SQLiteCommand command = new SQLiteCommand(connection))
            {
                command.CommandText = $"UPDATE [{tableName}] SET {today}='{hours}' WHERE name=@name";
                command.Parameters.AddWithValue("@name", student);
                command.Parameters.AddWithValue($"@{today}", hours);
                command.ExecuteNonQuery();
            }
        }
        private static void AddStudentNullWithDay(string tableName, SQLiteConnection connection, string hours, string student, string day)
        {
            string today = $"day{day}";
            SQLiteCommand add = new SQLiteCommand($"UPDATE [{tableName}] SET {today}=@hours WHERE name=@name", connection);
            add.Parameters.AddWithValue("@name", student);
            add.Parameters.AddWithValue("@hours", hours);
            add.ExecuteNonQuery();
        }
        /*public static void CheckAndCreateMonthTable()
        {
            using (SQLiteConnection sqlite_conn = new SQLiteConnection(database.connection))
            {
                sqlite_conn.Open();
                using (SQLiteCommand sqlite_cmd = sqlite_conn.CreateCommand())
                {
                    string currentMonth = DateTime.Now.ToString("MMMM", new System.Globalization.CultureInfo("en-US"));
                    string monthTable = currentMonth.ToLower();

                    sqlite_cmd.CommandText = $"SELECT count(*) FROM sqlite_master WHERE type='table' AND name='{monthTable}'";
                    int tableCount = Convert.ToInt32(sqlite_cmd.ExecuteScalar());

                    if (tableCount == 0)
                    {
                        sqlite_cmd.CommandText = $"CREATE TABLE {monthTable} (id INTEGER PRIMARY KEY, name TEXT, day1 INTEGER, day2 INTEGER, day3 INTEGER, day4 INTEGER, day5 INTEGER, day6 INTEGER, day7 INTEGER, day8 INTEGER, day9 INTEGER, day10 INTEGER, day11 INTEGER, day12 INTEGER, day13 INTEGER, day14 INTEGER, day15 INTEGER, day16 INTEGER, day17 INTEGER, day18 INTEGER, day19 INTEGER, day20 INTEGER, day21 INTEGER, day22 INTEGER, day23 INTEGER, day24 INTEGER, day25 INTEGER, day26 INTEGER, day27 INTEGER, day28 INTEGER, day29 INTEGER, day30 INTEGER, day31 INTEGER)";
                        sqlite_cmd.ExecuteNonQuery();
                    }
                }
                sqlite_conn.Close();
            }
        }*/
        #endregion
    
    }
}

    
