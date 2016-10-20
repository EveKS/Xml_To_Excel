using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xml_To_Excel.Presenter;
using Xml_To_Excel.Services;
using Xml_To_Excel.Utility;

namespace Xml_To_Excel
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            MainForm form = new MainForm();
            MessageService service = new MessageService();
            FileManager manager = new FileManager();

            MainPresenter presenter = new MainPresenter(form, manager, service);

            Application.Run(form);
        }
    }
}
