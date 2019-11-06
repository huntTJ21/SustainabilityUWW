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
        #region Classes

        #region Auth Setting
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
        private enum AuthTypes
        {
            [Description("Windows Authentication")]
            Windows,
            [Description("SQL Server Authentication")]
            Sqlserv
        }
        #endregion

        class SettingsPropertyGrid
        {
            #region Member Declarations
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
            
            [Browsable(true)]
            [Description("Location to store temporary Excel files.")]
            [Category("Excel")]
            [DisplayName("Temp File Path")]
            public string D_TempPath { get; set; } = "TempPath";
            #endregion

            #region Functions
            public string this[string key]
            {
                // this function defines a [] operator for this class
                get
                {
                    switch (key)
                    {
                        case "ServerURL": return D_ServerURL;
                        case "DBName": return D_DBName;
                        case "Authentication": return Enum.GetName(typeof(AuthTypes), D_Authentication);
                        case "TempPath": return D_TempPath;
                        default: throw new KeyNotFoundException();
                    }
                }
                set
                {
                    switch (key)
                    {
                        case "ServerURL": D_ServerURL = value.ToString();
                            break;
                        case "DBName": D_DBName = value.ToString();
                            break;
                        case "Authentication": Enum.TryParse(value.ToString(), out AuthTypes temp);
                            D_Authentication = temp;
                            break;
                        case "TempPath": D_TempPath = value.ToString();
                            break;
                        default: throw new KeyNotFoundException();
                    }
                }
            }
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

                #region JSON Handling
            public bool saveJSON(string path)
            {
                try
                {
                    // Create new JObject to hold the settings reflected 
                    //  in the propertygrid to serialize into a JSON file
                    dynamic saveObj = new JObject();
                    saveObj.ServerURL = D_ServerURL;
                    saveObj.Authentication = Enum.GetName(typeof(AuthTypes), D_Authentication); // Have to convert the ENUM value (0 or 1) to its name for storage
                    saveObj.DBName = D_DBName;
                    saveObj.TempPath = D_TempPath;

                    // Write generated JSON to file specified by path variable
                    File.WriteAllText(path, saveObj.ToString());
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return false;
                }
            }
            
            public bool verifyJSON(JObject json)
            {
                // Verify that all of the necessary keys are present in the JSON file.
                bool isCorrupt = false;
                if (!json.ContainsKey("ServerURL")){ isCorrupt = true; }
                if (!json.ContainsKey("Authentication")){ isCorrupt = true; }
                if (!json.ContainsKey("DBName")){ isCorrupt = true; }
                if (!json.ContainsKey("TempPath")) { isCorrupt = true; }
                return !isCorrupt;
            }
            public bool loadJSON(string path)
            {
                try
                {
                    // Load and parse JSON from the settings file. Then verif that all settings are present and that the file is not corrupt.
                    JObject settingsJSON = JObject.Parse(File.ReadAllText(path));
                    if (!verifyJSON(settingsJSON)) { throw new Exception("Some settings are not present. Settings file corrupt."); }

                    // Parse out the values to this class's members
                    D_ServerURL = settingsJSON.Value<String>("ServerURL");
                    D_DBName = settingsJSON.Value<String>("DBName");
                    Enum.TryParse(settingsJSON.Value<String>("Authentication"), out AuthTypes authTemp);
                    D_Authentication = authTemp;
                    D_TempPath = settingsJSON.Value<String>("TempPath");

                    return true;
                }
                catch (FileNotFoundException)
                {
                    // The settings file does not exist, create a new one
                    MessageBox.Show("Settings file could not be found.\nA new settings file will be created.", "Error Loading Settings");
                    // This object is constructed with default values, so just use the saveJSON() method to save the default values.
                    return saveJSON(path);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.GetType().ToString());
                    // Get the location of the current settings file
                    String dir = Path.GetDirectoryName(path);
                    // Create a new path string with the same directory as the current settings file, but with the filename 'settings_old.json'
                    String newPath = Path.Combine(dir, "settings_old.json");

                    // If the file cannot be read for any reason, inform the user that it will be recreated.
                    // Since we know that a settings.json file was detected (otherwise the first exception would have triggered),
                    //      We should rename it and inform the user so that they can retrieve corrupted data if needed.
                    MessageBox.Show("Settings file could not be loaded.\n\nError:\n"+ ex.Message + "\n\nA new settings file will be created.\nThe corrupted file will be moved to the path below so that your settings may be retrieved by hand.\n\n" + Path.GetFullPath(newPath), "Error loading settings");

                    // Use the File.Move method to efectively 'rename' the settings file.
                    if (File.Exists(newPath))
                    {
                        File.Delete(newPath);
                    }
                    File.Move(path, newPath);

                    // This object is constructed with default values, so just use the saveJSON() method to save the default values.
                    return saveJSON(path);
                }
            }
                #endregion

            #endregion
        }
        #endregion

        #region Variables
        SettingsPropertyGrid spg;
        String settingsFilePath = @".\settings.json";
        #endregion

        #region Constructor
        public win_Settings()
        {
            spg = new SettingsPropertyGrid();
            spg.loadJSON(settingsFilePath);
            InitializeComponent();
            prop_Settings.SelectedObject = spg;
        }
        #endregion

        #region Functions
        /*
        public string DBName
        {
            get
            {
                return spg["DBName"];
            }
            set
            {
                spg["DBName"] = value.ToString();
            }
        }
        public string ServerURL
        {
            get
            {
                return spg["DBName"];
            }
            set
            {
                spg["DBName"] = value.ToString();
            }
        }
        */
        #endregion

        #region Event Listeners
        private void btn_TestConnection_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("DBName: " + spg["DBName"] + "\nServerURL: " + spg["ServerURL"]);
            DBConnection db = new DBConnection(spg["ServerURL"], spg["DBName"]);
            MessageBox.Show(db.Connect()
               ? "Connection Successful!"
               : "Connection Unsuccessful. Please verify settings and try again.");
            if (db.IsConnected())
                db.Disconnect();
        }

        private void prop_Settings_PropertyValueChanged(object sender, PropertyValueChangedEventArgs e)
        {
            var props = prop_Settings.Properties;
            foreach(PropertyItem prop in props){
                String propName = SettingsPropertyGrid.Convert(prop.DisplayName);
                spg[propName] = prop.Value.ToString();
            }
        }

        private void btn_SaveAndExit_Click(object sender, RoutedEventArgs e)
        {
            // When the user hits this button, verify all settings and save them to the settings file. Then close the settings window.
            if (File.Exists(settingsFilePath))
            {
                File.Delete(settingsFilePath);
            }
            spg.saveJSON(settingsFilePath);
            this.Close();
        }
        #endregion
    }
}
