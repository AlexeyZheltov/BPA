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
        
        public bool GetSetting(BPASettingEnum typeSetting, out (string BudgetPath, string DecisionPath, string PriceListMTPath) settings_var)
        {
            if (RequestSettings(typeSetting) == DialogResult.OK)
            {
                string budget_path = (string)settings["BudgetPath"];
                string decision_path = (string)settings["DecisionPath"];
                string price_path = (string)settings["PriceListMTPath"];

                settings_var = (File.Exists(budget_path) ? budget_path : "", File.Exists(decision_path) ? decision_path : "", File.Exists(price_path) ? price_path : "");
                return true;
            }
            else
            {
                settings_var = ("", "", "");
                return false;
            }
            
        }

        public bool GetBudgetPath(out string path, bool requestFile = false)
        {
            var (_result, _path) = GetSetting("BudgetPath", requestFile, BPASettingEnum.Budget);
            path = _path;
            return _result;
        }

        public bool GetDecisionPath(out string path, bool requestFile = false)
        {
            var (_result, _path) = GetSetting("DecisionPath", requestFile, BPASettingEnum.Decision);
            path = _path;
            return _result;
        }

        public bool GetPriceListMT(out string path, bool requestFile = false)
        {
            var (_result, _path) = GetSetting("PriceListMTPath", requestFile, BPASettingEnum.PriceListMT);
            path = _path;
            return _result;
        }

        private (bool Status, string Path) GetSetting(string prop_name, bool requestFile, BPASettingEnum typeSetting)
        {
            if (!requestFile && File.Exists((string)settings[prop_name]))
                return (true, (string)settings[prop_name]);
            
            if (RequestSettings(typeSetting) == DialogResult.OK && File.Exists((string)settings[prop_name]))
            {
                return (true, (string)settings[prop_name]);
            }
            else return (false, "");
        }

        private DialogResult RequestSettings(BPASettingEnum typeSetting)
        {
            DialogResult result;
            Forms.SettingsForm settingsForm = new Forms.SettingsForm(typeSetting);
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

    public enum BPASettingEnum
    {
        Budget = 1,
        Decision = 2,
        PriceListMT = 4,
        All = 7
    }
}
