using System.Windows.Forms;

namespace Xml_To_Excel.Services
{
    public interface IMessageService
    {
        void ShowError(string error);
        void ShowExclamation(string exlamation);
        void ShowMessage(string message);
    }

    public class MessageService : IMessageService
    {
        public void ShowMessage(string message)
        {
            MessageBox.Show(message, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ShowExclamation(string exlamation)
        {
            MessageBox.Show(exlamation, "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        public void ShowError(string error)
        {
            MessageBox.Show(error, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
