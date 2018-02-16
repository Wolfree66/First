using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References.Users;

namespace WpfApp_DialogueAddSignatories
{
    class Presenter
    {
        Model model { get; set; }
        MainWindow mainWindow { get; set; }
        // Выбор_пользователя dialogUserSelection { get; set; }
        public Presenter(ServerConnection connection, MainWindow mainWindow, TFlex.DOCs.Model.References.ReferenceObject currentTechnicalSolution, string nameServer = "TFLEX")
        {
            if (this.mainWindow == null)
                this.mainWindow = mainWindow;


            if (model == null)
                this.model = new Model(connection, currentTechnicalSolution, nameServer);

            this.mainWindow.FindEmployeeAtPostEvent += mainWindow_FindEmployeeAtPostEvent;
            this.mainWindow.GetAllUserEvent += mainWindow_GetAllUsersEvent;
            this.mainWindow.SelectProgramManagerEvent += mainWindow_SelectProgramManagerEvent;
            this.mainWindow.OKEvent += mainWindow_OKEvent;
            this.mainWindow.ClearComboBoxEvent += mainWindow_ClearComboBoxEvent;
            this.mainWindow.SelectEmployeeEvent += mainWindow_SelectEmployeeEvent;
        }

        private void mainWindow_ClearComboBoxEvent(object sender, EventArgs e)
        {
            model.ProgramManager = null;
            mainWindow.comboBox.SelectedIndex = -1;
        }

        public Presenter(ServerConnection connection, MainWindow mainWindow, string nameServer = "TFLEX")
        {

            if (this.mainWindow == null)
                this.mainWindow = mainWindow;

            if (model == null)
                model = new Model(connection, nameServer);

            this.mainWindow.FindEmployeeAtPostEvent += mainWindow_FindEmployeeAtPostEvent;
            this.mainWindow.GetAllUserEvent += mainWindow_GetAllUsersEvent;
            this.mainWindow.SelectProgramManagerEvent += mainWindow_SelectProgramManagerEvent;
            this.mainWindow.OKEvent += mainWindow_OKEvent;
            this.mainWindow.ClearComboBoxEvent += mainWindow_ClearComboBoxEvent;
            this.mainWindow.SelectEmployeeEvent += mainWindow_SelectEmployeeEvent;
        }

        private void mainWindow_SelectEmployeeEvent(object sender, EventArgs e)
        {
            model.ChiefDesignerDirection = mainWindow.listBox_users.SelectedItem as User;
        }

        private void mainWindow_SelectProgramManagerEvent(object sender, EventArgs e)
        {
            User SelectedEmpl = this.mainWindow.comboBox.SelectedItem as User;

            if (SelectedEmpl == null) return;

            if (SelectedEmpl != null)
            {
                model.ProgramManager = SelectedEmpl;
            }
        }

        /* private void dialogUserSelection_SelectEmployeeEvent(object sender, EventArgs e)
         {
             model.ChiefDesignerDirection = dialogUserSelection.lstBox.SelectedItem as User;
         }
         private void dialogUserSelection_OKEvent(object sender, EventArgs e)
         {
             dialogUserSelection.Close();
         }*/
        private void mainWindow_OKEvent(object sender, EventArgs e)
        {
            if (model.ChiefDesignerDirection == null)
            {
                model.ShowErrorMsg("Не выбран главный конструктор по направлению!",
"Выберите главного конструктора по направлению");
                return;
            }
            try
            {
                bool resilt = model.AddSign();

                if (resilt)
                    System.Windows.Forms.MessageBox.Show("Подписи успешно добавлены");

                mainWindow.Close();

            }
            catch (Exception ex)
            {
                model.ShowErrorMsg(ex.Message,
"Ошибка!");
            }
            mainWindow.Close();
        }



        /// <summary>
        /// выбор должности
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void mainWindow_FindEmployeeAtPostEvent(object sender, EventArgs e)
        {
            object SelectedPost = this.mainWindow.listBox_posts.SelectedItem;
           
            if (SelectedPost == null)
            {
                System.Windows.MessageBox.Show("не выбрана должность");

                model.ShowErrorMsg("Выберите главного конструктора по направлению!",
    "Не выбрана дожность главного коструктора.");

            }
            else
            {

                string value = SelectedPost.ToString();

                List<User> UsersWorkOnPost = this.model.ExtensionDinamika.НайтиПользоватейПоДолжности(value);

                if (!model.UsersWithEquallyPost.ContainsKey(value))
                    model.add_UsersWithEquallyPost(value, UsersWorkOnPost);

                int count = model.UsersWithEquallyPost[value].Count;

                if (count == 0)
                {
                    model.ShowErrorMsg("По выбранной Вами должности не работают, выберите другую должность",
      "Ошибка!");
                    return;
                }
                if (count > 1)
                {
                    mainWindow.label2.Content = "По выбранной Вами должности работают несколько пользователей, выберите нужного";
                    mainWindow.label2.Visibility = Visibility.Visible;
                    mainWindow.listBox_users.Visibility = Visibility.Visible;
                    mainWindow.listBox_users.ItemsSource = model.UsersWithEquallyPost[value];
                }
                else if (count == 1)
                {
                    mainWindow.label2.Visibility = Visibility.Visible;
                    mainWindow.listBox_users.Visibility = Visibility.Visible;
                    mainWindow.label2.Content = "Работник по выбранной должности";
                    mainWindow.listBox_users.ItemsSource = model.UsersWithEquallyPost[value];
                    mainWindow.listBox_users.SelectedItem = mainWindow.listBox_users.Items[0];
                    // model.ChiefDesignerDirection = mainWindow.listBox_users.SelectedItem as User;
                }
            }
        }

        /// <summary>
        /// Загрузка окна
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void mainWindow_GetAllUsersEvent(object sender, EventArgs e)
        {
            model.AllUsers = this.model.ExtensionDinamika.ПолучитьВсехПользователейОрганизационнойСтруктуры();
            foreach (var DataUser in model.AllUsers)
            {
                mainWindow.comboBox.Items.Add(DataUser);
            }

        }
    }
}