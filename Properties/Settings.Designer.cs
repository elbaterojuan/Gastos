﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Gastos.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.3.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(@"<?xml version=""1.0"" encoding=""utf-16""?>
<ArrayOfString xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
  <string>Supermercado</string>
  <string>Verduleria</string>
  <string>Almacen</string>
  <string>Carnicería</string>
  <string>Pescadería</string>
  <string>Sushi</string>
  <string>Expensas</string>
  <string>Edesur</string>
  <string>Metrogas</string>
  <string>Impuesto municipal</string>
  <string>Arba</string>
  <string>Telefónica</string>
  <string>Casamiento</string>
  <string>Netflix</string>
  <string>Cablevisión</string>
  <string>Prestamo</string>
  <string>Salida</string>
  <string>Seguro Hogar</string>
  <string>Spotify</string>
  <string>Farmacia</string>
  <string>Limpieza</string>
  <string>Consolidación</string>
  <string>Otros</string>
</ArrayOfString>")]
        public global::System.Collections.Specialized.StringCollection Categorias {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["Categorias"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("C:\\Users\\juan.sebastian.rocco\\Downloads")]
        public string Carpeta {
            get {
                return ((string)(this["Carpeta"]));
            }
            set {
                this["Carpeta"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Gastos.xls")]
        public string Archivo {
            get {
                return ((string)(this["Archivo"]));
            }
            set {
                this["Archivo"] = value;
            }
        }
    }
}
