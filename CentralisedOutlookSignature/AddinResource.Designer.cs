﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.34209
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CentralisedOutlookSignature {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class AddinResource {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal AddinResource() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("CentralisedOutlookSignature.AddinResource", typeof(AddinResource).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Signatur aktualisiert.
        /// </summary>
        internal static string Addin_NewSigHeader {
            get {
                return ResourceManager.GetString("Addin_NewSigHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Die Signatur wurde aktualisiert. Damit die Änderung wirksam wird, muss Outlook neu gestartet werden. Outlook jetzt schließen?.
        /// </summary>
        internal static string Addin_NewSigText {
            get {
                return ResourceManager.GetString("Addin_NewSigText", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to http://schemas.microsoft.com/mapi/proptag/0x661D000B.
        /// </summary>
        internal static string schemaOoO {
            get {
                return ResourceManager.GetString("schemaOoO", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Signaturen aktualisiert.
        /// </summary>
        internal static string SignaturesUpdatedHeader {
            get {
                return ResourceManager.GetString("SignaturesUpdatedHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Die Signaturen wurden aktualisiert..
        /// </summary>
        internal static string SignaturesUpdatedText {
            get {
                return ResourceManager.GetString("SignaturesUpdatedText", resourceCulture);
            }
        }
    }
}
