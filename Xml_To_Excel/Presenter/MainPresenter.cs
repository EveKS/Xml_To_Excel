using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xml_To_Excel.Services;
using Xml_To_Excel.Utility;

namespace Xml_To_Excel.Presenter
{
    public class MainPresenter
    {
        private readonly IMainForm _view;
        private readonly IFileManager _meneger;
        private readonly IMessageService _messegeService;

        private string _currentExcelFilePath;
        private string _currentXmlFolderPath;

        public MainPresenter(IMainForm view, IFileManager manager, IMessageService service)
        {
            _view = view;
            _meneger = manager;
            _messegeService = service;

            _view.SelectExel += _view_SelectExel;
            _view.SelectXml += _view_SelectXml;
            _view.Save += _view_Save;
        }

        private void _view_SelectExel(object sender, EventArgs e)
        {
            try
            {
                string exelPath = _view.SelectExelPath;

                bool isExist = _meneger.IsExist(exelPath);

                if (!isExist)
                {
                    _messegeService.ShowExclamation("Выбранный фаил не существует.");
                    return;
                }

                _currentExcelFilePath = exelPath;
            }
            catch (Exception ex)
            {
                _messegeService.ShowError(ex.Message);
            }
        }
        private void _view_SelectXml(object sender, EventArgs e)
        {
            try
            {
                string xmlsPath = _view.SelectXmlFolderPath;

                bool isFolderExist = _meneger.IsFolderExist(xmlsPath);

                if (!isFolderExist)
                {
                    _messegeService.ShowExclamation("Выбранная папка не существует.");
                    return;
                }

                _currentXmlFolderPath = xmlsPath;
            }
            catch (Exception ex)
            {
                _messegeService.ShowError(ex.Message);
            }
        }
        private async void _view_Save(object sender, EventArgs e)
        {
            try
            {
                await _meneger.GetExcel(_currentExcelFilePath, _currentXmlFolderPath);
            }
            catch (Exception ex)
            {
                _messegeService.ShowError(ex.Message);
            }
        }        
    }
}
