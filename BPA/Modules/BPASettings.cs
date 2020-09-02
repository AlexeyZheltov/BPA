using BPA.Properties;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BPA.Modules
{
    /// <summary>
    /// Файл обертка над ндстройками
    /// </summary>
    class BPASettings
    {
        readonly Settings settings = Properties.Settings.Default;

        public bool GetBudgetPath(out string path, bool requestFile = false)
        {
            var (_result, _path) = GetSetting("BudgetPath", requestFile);
            path = _path;
            return _result;
        }

        public bool GetDecisionPath(out string path, bool requestFile = false)
        {
            var (_result, _path) = GetSetting("DecisionPath", requestFile);
            path = _path;
            return _result;
        }

        public bool GetPriceListMT(out string path, bool requestFile = false)
        {
            var (_result, _path) = GetSetting("PriceListMTPath", requestFile);
            path = _path;
            return _result;
        }

        private (bool, string) GetSetting(string prop_name, bool requestFile)
        {
            if (!requestFile && File.Exists((string)settings[prop_name]))
                return (true, (string)settings[prop_name]);
            
            if (RequestSettings() == DialogResult.OK && File.Exists((string)settings[prop_name]))
            {
                return (true, (string)settings[prop_name]);
            }
            else return (false, "");
        }

        private DialogResult RequestSettings()
        {
            DialogResult result;
            Forms.SettingsForm settingsForm = new Forms.SettingsForm();
            result = settingsForm.ShowDialog(new ExcelWindows(Globals.ThisWorkbook));
            settingsForm.Close();
            return result;
        }

        public enum GetSettingResultEnum
        {
            Ok,
            Cancel
        }
    }
}
