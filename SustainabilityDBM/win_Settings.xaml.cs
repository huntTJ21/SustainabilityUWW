using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Xceed.Wpf.Toolkit.PropertyGrid;
using Path = System.IO.Path;

namespace SustainabilityDBM
{
    /// <summary>
    /// Interaction logic for win_Settings.xaml
    /// </summary>
    public partial class win_Settings : Window
    {
        // Classes
        class AuthConverter : EnumConverter
        {
            private Type _enumType;

            public AuthConverter(Type type)
                : base(type)
            {
                _enumType = type;
            }

            public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
            {
                return destinationType == typeof(string);
            }

            public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture,
                object value, Type destType)
            {
                FieldInfo fi = _enumType.GetField(Enum.GetName(_enumType, value));
                DescriptionAttribute dna =
                    (DescriptionAttribute)Attribute.GetCustomAttribute(fi, typeof(DescriptionAttribute));

                if (dna != null)
                    return dna.Description;
                else
                    return value.ToString();
            }

            public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
            {
                return sourceType == typeof(string);
            }

            public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
            {
                foreach (FieldInfo fi in _enumType.GetFields())
                {
                    DescriptionAttribute dna =
                        (DescriptionAttribute)Attribute.GetCustomAttribute(fi, typeof(DescriptionAttribute));

                    if ((dna != null) && ((string)value == dna.Description))
                        return Enum.Parse(_enumType, fi.Name);
                }

                return Enum.Parse(_enumType, (string)value);
            }
        }

        class SettingsPropertyGrid
        {
            // Assume that all data is already present

            [Browsable(true)]
            [Description("Machine/Name of SQL Server Instance")]
            [Category("Database Connection")]
            [DisplayName("Server URL")]
            public string D_ServerURL { get; set; } = "ServerURL";

            [Browsable(true)]
            [Description("Authentication to use for SQL Server Connection")]
            [Category("Database Connection")]
            [DisplayName("Authentication Type")]
            [TypeConverter(typeof(AuthConverter))]
            public AuthTypes D_Authentication { get; set; }

            [Browsable(true)]
            [Description("Name of Database to use on the server.")]
            [Category("Database Connection")]
            [DisplayName("Database Name")]
            public string D_DBName { get; set; } = "DBName";

            public static string Convert(string display)
            {
                switch (display)
                {
                    case "Database Name": return "DBName";
                    case "Authentication Type": return "Authentication";
                    case "Server URL": return "ServerURL";
                    default: throw new Exception("Invalid Field");
                }
            }
        }

        private enum AuthTypes
        {
            [Description("Windows Authentication")]
            Windows,
            [Description("SQL Server Authentication")]
            Sqlserv
        }

        // Variables
        Dictionary<string, string> dict_settings = new Dictionary<string, string>();
        String test = "";
        // Constructor
        public win_Settings()
        {
            var spg = new SettingsPropertyGrid();
            InitializeComponent();
            dict_settings["Authentication"] = "Windows";
            dict_settings["ServerURL"] = "ServerURL";
            dict_settings["DBName"] = "DevStaging";
            prop_Settings.SelectedObject = spg;
            loadSettingsFromJSON();
        }

        // Functions
        public string SettingsUri
        {
            get
            {
                MessageBox.Show(Path.GetFullPath("settings.json"));
                return Path.GetFullPath("settings.json");
            }
        }
        private bool loadSettingsFromJSON()
        {
            // TODO: Implement loadSettingsFromJSON()
            // Testing Commits
            try
            {
                //var ofd = new Microsoft.Win32.OpenFileDialog();
                //var result = ofd.ShowDialog();
                //if (result == false) return false;
                //MessageBox.Show(ofd.FileName);
                string test = @"{
                    Connection: {
                        ServerURL: 'Test'
                    }
                }";
                // var o1 = JObject.Parse(File.ReadAllText(SettingsUri));
                using (StreamReader file = File.OpenText(@".\settings.json"))
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject o2 = (JObject)JToken.ReadFrom(reader);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

        }

        // Event Listeners
        private void btn_TestConnection_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("DBName: " + dict_settings["DBName"] + "\nServerURL: " + dict_settings["ServerURL"]);
            DBConnection db = new DBConnection(dict_settings["ServerURL"], dict_settings["DBName"]);
            MessageBox.Show(db.Connect()
               ? "Connection Successful!"
               : "Connection Unsuccessful. Please verify settings and try again.");
            if (db.IsConnected())
                db.Disconnect();
        }

        private void prop_Settings_PropertyValueChanged(object sender, PropertyValueChangedEventArgs e)
        {
            //MessageBox.Show("PropertyValue\n" + e.NewValue);
            var props = prop_Settings.Properties;
            foreach(PropertyItem prop in props){
                String propName = SettingsPropertyGrid.Convert(prop.DisplayName);
                dict_settings[propName] = prop.Value.ToString();
            }
            //dict_settings[SettingsPropertyGrid.Convert(e.ChangedItem.Label)] = e.ChangedItem.Value.ToString();
            //MessageBox.Show(dict_settings[SettingsPropertyGrid.Convert(e.ChangedItem.Label)]);
        }
    }
}
