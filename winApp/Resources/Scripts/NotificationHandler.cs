using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace winApp.Resources
{
    class NotificationHandler
    {
        LoadingWindow loadingWindow;
        MessageWindow messageWindow;

        public NotificationHandler(System.Windows.Window owner)
        {
            loadingWindow = new LoadingWindow();
            messageWindow = new MessageWindow();
            loadingWindow.Owner = owner;
            messageWindow.Owner = owner;
        }

        public void Message(int code)
        {
            loadingWindow.CloseNote();
            string message;
            switch (code)
            {
                case 0:
                    message = "Успешно!";
                    break;
                case 1:
                    message = "Недопустимая сумма!";
                    break;
                case 2:
                    message = "Не введена сумма!";
                    break;
                case 3:
                    message = "Не выбран радел!";
                    break;
                default:
                    message = "This case is not exist";
                    break;
            }
            messageWindow.ShowMessage(message);
        }
        public void Note(int code)
        {
            string note;
            switch (code)
            {
                case 0:
                    note = "Пожалуйста, подождите...";
                    break;
                case 1:
                    note = " ";
                    break;
                case 2:
                    note = " ";
                    break;
                case 3:
                    note = " ";
                    break;
                default:
                    note = "This case is not exist";
                    break;
            }
            loadingWindow.ShowNote(note);
        }
    }
}
