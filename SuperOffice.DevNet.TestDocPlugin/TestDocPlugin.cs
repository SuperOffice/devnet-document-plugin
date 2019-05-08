using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using SuperOffice.CRM.Documents;
using System.IO;

namespace SuperOffice.DevNet.TestDocPlugin
{
	/// <summary>
	/// This is a test and example of how to implement a document plugin.
	/// Documents are stored in a folder specified in the CONFIG file.
	/// <para/>
	/// This Test Doc Plugin has doc plugin id = 123.
	/// <para/>
	/// The plugin creates files in "C:\TestDoc\Documents" named "glops docid-123.docx".
	/// Templates are stored in "C:\TestDoc\Templates" folder.
	/// <para/>
	/// Checking in and out are handled using a ".lock" file that contains the name of the user who locked the file.
	/// Delete the file to revert the lock.
	/// <para/>
	/// Versions are created when checking in, by making  copies of the file as "glops docid-123 v1.docx" and "glops docid-123 v2.docx" and so on.
	/// </summary>
	/// <example>
	/// <code>
	///  &lt;SuperOffice&gt;
	///    &lt;Documents&gt;
	///     &lt;add key="TestDocPlugin.Path" value="c:\testDoc" /&gt;
	///     &lt;add key="TestDocPlugin.CanCreateDocumentTemplates" value="true" /&gt;
	///     &lt;add key="TestDocPlugin.CanLock" value="true" /&gt;
	///     &lt;add key="TestDocPlugin.CanVersion" value="true" /&gt;
	///     &lt;add key="TestDocPlugin.CanCommands" value="true" /&gt;
	///    &lt;/Documents&gt;
	///  &lt;/SuperOffice&gt;
	/// </code>
	/// </example>
	[DocumentPlugin2("Test Doc Plugin", 123)]
	public class TestDocPlugin : SuperOffice.CRM.Documents.IDocumentPlugin2
	{
		private SuperOffice.CRM.IConfiguration _config;

        /// <summary>
        /// Determines if this plugin store multiple versions of documents
        /// </summary>
		public static bool OverrideCanVersion = false;

        /// <summary>
        /// Determines if this plugin can lock documents
        /// </summary>
        public static bool OverrideCanLock = false;

        /// <summary>
        /// Determines if this plugin can invoke commands
        /// </summary>
        public static bool OverrideCanCommands = false;

		/// <summary>
		/// Constructor - receives IConfiguration to read configuration settings.
		/// </summary>
		public TestDocPlugin(SuperOffice.CRM.IConfiguration config)
		{
			Log("TestDocPlugin ctor");
			_config = config;
			Log("TestDocPlugin rootpath = " + RootPath);
			EnsurePathsExist();
			Log("TestDocPlugin ctor done");
		}

		#region Config Settings

		/// <summary>
		/// Root path for archives and templates
		/// </summary>
		/// <example>
		/// <code>
		///  &lt;SuperOffice&gt;
		///    &lt;Documents&gt;
		///     &lt;add key="TestDocPlugin.Path" value="c:\testDoc" /&gt;
		///    &lt;/Documents&gt;
		///  &lt;/SuperOffice&gt;
		/// </code>
		/// </example>
		public string RootPath
		{
			get
			{
				string s = _config.GetConfigString("SuperOffice/Documents/TestDocPlugin.Path") ?? "c:\\testDoc";
				return s;
			}
		}

		/// <summary>
		/// Can this plugin create document templates?
		/// </summary>
		/// <example>
		/// <code>
		///  &lt;SuperOffice&gt;
		///    &lt;Documents&gt;
		///     &lt;add key="TestDocPlugin.CanCreateDocumentTemplates" value="true" /&gt;
		///    &lt;/Documents&gt;
		///  &lt;/SuperOffice&gt;
		/// </code>
		/// </example>
		public bool CanCreateDocumentTemplates
		{
			get
			{
				bool res = false;
				res = _config.GetConfigBool("SuperOffice/Documents/TestDocPlugin.CanCreateDocumentTemplates");
				return res;
			}
		}

		/// <summary>
		/// Can this plugin store multiple versions of documents?
		/// </summary>
		/// <example>
		/// <code>
		///  &lt;SuperOffice&gt;
		///    &lt;Documents&gt;
		///     &lt;add key="TestDocPlugin.CanVersion" value="true" /&gt;
		///    &lt;/Documents&gt;
		///  &lt;/SuperOffice&gt;
		/// </code>
		/// </example>
		public bool CanVersion
		{
			get
			{
				bool res = false;
				res = _config.GetConfigBool("SuperOffice/Documents/TestDocPlugin.CanVersion");
				return res || OverrideCanVersion;
			}
		}


		/// <summary>
		/// Can this plugin lock documents (check-in and check-out)?
		/// </summary>
		/// <example>
		/// <code>
		///  &lt;SuperOffice&gt;
		///    &lt;Documents&gt;
		///     &lt;add key="TestDocPlugin.CanLock" value="true" /&gt;
		///    &lt;/Documents&gt;
		///  &lt;/SuperOffice&gt;
		/// </code>
		/// </example>
		public bool CanLock
		{
			get
			{
				bool res = false;
				res = _config.GetConfigBool("SuperOffice/Documents/TestDocPlugin.CanVersion");
				return res || OverrideCanLock;
			}
		}


		/// <summary>
		/// Can this plugin Add custom commands?
		/// </summary>
		/// <example>
		/// <code>
		///  &lt;SuperOffice&gt;
		///    &lt;Documents&gt;
		///      &lt;add key="TestDocPlugin.CanCommands" value="true" /&gt;
		///    &lt;/Documents&gt;
		///  &lt;/SuperOffice&gt;
		/// </code>
		/// </example>
		public bool CanCommands
		{
			get
			{
				bool res = false;
				res = _config.GetConfigBool("SuperOffice/Documents/TestDocPlugin.CanCommands");
				//string s = ConfigurationManager.AppSettings.Get("TestDocPlugin.CanCommands") ?? "";
				//bool.TryParse(s, out res);
				return res || OverrideCanCommands;
			}
		}

        #endregion


        /// <summary>
        /// Check in a currently checked-out document
        /// </summary>
        /// <remarks>
        /// If the document plugin supports locking and the requesting user is the one who checked out the document, 
        /// then the last-saved content by that user should become the new publicly visible content, and 
        /// the checkout state should be reset. Calls by other users should result in failure and no state change.
        /// <para/>
        /// If the document plugin does not support locking or versioning, then this call should perform no action.
        /// </remarks>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type. ("None", "Message", "SoProtocol", "CustomGUI", "Other")</param>
        /// <param name="versionDescription">Version description.</param>
        /// <param name="versionExtraFields">Extra fields</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
        public CRM.ReturnInfo CheckinDocument(CRM.IDocumentInfo documentInfo, string[] allowedReturnTypes, string versionDescription, string[] versionExtraFields)
        {
            Log("TestDocPlugin CheckinDocument");
            var res = new CRM.ReturnInfo() { Success = false };
            if (CanLock)
            {
                string currentUser = SoContext.CurrentPrincipal.Associate;
                string path = GetPath(documentInfo);
                string locker = GetCheckoutInfo(path);
                if (locker == "")
                {
                    res.Success = false;
                    res.Type = CRM.ReturnType.Message;
                    res.Value = "Not checked out";
                }
                else
                    if (locker == currentUser)
                    res.Success = true;
                else
                {
                    res.Success = false;
                    res.Type = CRM.ReturnType.Message;
                    res.Value = "Checked out to " + locker;
                }

                if (res.Success)
                {
                    // find next available ver number
                    int verNum = 1;
                    string verPath = GetPath(documentInfo, verNum.ToString());
                    while (File.Exists(verPath))
                    {
                        verNum++;
                        verPath = GetPath(documentInfo, verNum.ToString());
                    }
                    File.Copy(path, verPath);
                    SetCheckoutInfo(path, "");

                }
            }
            return res;
        }

        /// <summary>
        /// Check out the document for editing
        /// </summary>
        /// <remarks>
        /// A document plugin that supports versioning may internally prepare to receive new content and 
        /// prepare a new internal version, but a subsequent GetDocumentVersionList call should <b>not</b> 
        /// show this version – not until CheckInDocument has been called. 
        /// <para/>
        /// After the completion of this call, the document is in checked out state and <see cref="GetCheckoutState"/> 
        /// should return “Own” as the status. <see cref="SaveDocumentFromStream"/> calls on behalf of other users should
        /// fail, as should <see cref="CheckoutDocument"/> and <see cref="CheckinDocument"/> calls on behalf of other users.
        /// <para/>
        /// If the document plugin does not support locking or versioning, then this call should perform no action.
        /// </remarks>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type.</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
        public CRM.ReturnInfo CheckoutDocument(CRM.IDocumentInfo documentInfo, string[] allowedReturnTypes)
        {
            Log("TestDocPlugin CheckoutDocument");
            var res = new CRM.ReturnInfo() { Success = false };
            if (CanLock)
            {
                string currentUser = SoContext.CurrentPrincipal.Associate;
                string path = GetPath(documentInfo);
                string locker = GetCheckoutInfo(path);
                if (locker == "")
                    res.Success = true;
                else
                    if (locker == currentUser)
                    res.Success = true;
                else
                    res.Success = false;

                if (res.Success)
                    SetCheckoutInfo(path, currentUser);
                else
                {
                    res.Type = CRM.ReturnType.Message;
                    res.Value = "Checked out by " + locker;
                }
            }
            return res;
        }

        /// <summary>
        /// Create a default document based on the given documentType. Called when creating a new template.
        /// </summary>
        /// <param name="documentTypeKey">Id for a document type. NULL or blank if no types are supported.</param>
        /// <param name="documentTemplateInfo">Document template info</param>
        /// <returns>Extref/Filename for new template. This value is written to the template's Filename property.
        /// Return NULL if no change, or if no document created.</returns>
		public CRM.Documents.TemplateInfo CreateDefaultDocumentTemplate(int documentTypeKey, CRM.IDocumentTemplateInfo documentTemplateInfo)
        {
            Log("TestDocPlugin CreateDefaultDocumentTemplate");
            if (CanCreateDocumentTemplates)
            {
                string path = GetPath(documentTemplateInfo);        // arc\Templates
                string name = documentTemplateInfo.Name ?? "name";
                string ext = "";
                string mime = "";
                switch (documentTypeKey)
                {
                    case 1: ext = ".docx"; mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; break;
                    case 2: ext = ".txt"; mime = "text/plain"; break;
                    case 3: ext = ".jpg"; mime = "image/jpeg"; break;
                }

                var res = new TemplateInfo()
                {
                    MimeType = mime,
                    PluginId = 123,
                    Name = documentTemplateInfo.Name,
                    ExternalReference = Path.Combine(path, name + ext)
                };

                // Create an empty file
                File.Create(res.ExternalReference).Dispose();

                return res;
            }
            else
                return null;
        }

        /// <summary>
        /// Create a document in the repository, intially without content
        /// </summary>
        /// <remarks>
        /// Execution of this method should result in the creation of a document instance in the
        /// underlying repository, with empty content. If locking is supported, the status should
        /// be 'Checked-out'; the document should not be visible to other users. If locking is not 
        /// supported, a zero-length content sh.ould be the result.
        /// <para/>
        /// All metadata should be saved, an externalReference key should be assigned, and
        /// the file name/document name validated and reserved.
        /// <para/><b>Notes on semantics</b><br/>
        /// The IDocumentInfo.<see cref="SuperOffice.CRM.IDocumentInfo.LockSemantics"/> property is used to specify the
        /// locking/versioning semantics requested and implemented for a document. Semantics actually implemtned
        /// will be the lowest of what is requested and what is supported. Thus, a document may request 
        /// <see cref="SuperOffice.CRM.Documents.DocumentLockSemantics.None"/> semantics even if versioning is optionally supported
        /// by the document plugin, and in that case no versioning or locking should be performed.
        /// <para/>
        /// On creation, if locking and/or versioning is supported and requested, then the document stream should
        /// be saved to a temporary area. Calls to GetDocumentStream should return an empty stream until the first
        /// call to CheckinDocument has succeeded. The result of that Checkin call will be the base version 
        /// (version 1).
        /// <para/>
        /// Locking semantics are set on creation and cannot be changed later.
        /// </remarks>
        /// <param name="incomingInfo">SuperOffice metadata for the document, including the document Id
        /// and locking/versioning sematics requested for this document.</param>
        /// <param name="fileName">Suggested file name/document name. The document plugin must validate this
        /// name and amend it if needed (this is a ref parameter). If the name needs to be changed
        /// for any reason, a new and valid name must be generated by the plugin, and returned
        /// to the caller.</param>
        /// <param name="extraFields">Plugin-dependent metadata for the document as a whole. The
        /// usual caveats apply, i.e., there may be name/value pairs in the parameter that
        /// belong to other parts of the system. Failure to recognize a name is not an error.</param>
        /// <param name="versionDecription">Description of the initial version, if versioning is supported
        /// and enabled for the document.</param>
        /// <param name="versionExtraFields">Plugin-dependent metadata for the initial version</param>
        /// <returns>The external reference that identifies this document.</returns>
		public string CreateDocument(CRM.IDocumentInfo incomingInfo, ref string fileName, string[] extraFields, string versionDecription, string[] versionExtraFields)
        {
            Log("TestDocPlugin CreateDocument");
            if (!string.IsNullOrEmpty(incomingInfo.Name))
                fileName = incomingInfo.Name;
            string name = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);
            string path = GetPath(incomingInfo);
            string dir = Path.GetDirectoryName(path);

            // modify the filename stored on the document to ensure uniqueness
            fileName = name + " docid-" + incomingInfo.DocumentId.ToString() + ext;

            // return full path as extref
            string res = Path.Combine(dir, fileName);
            return res;
        }

        /// <summary>
        /// Delete a document, all versions and all metadata from the repository
        /// </summary>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type.</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
        public CRM.ReturnInfo DeleteDocument(CRM.IDocumentInfo documentInfo, string[] allowedReturnTypes)
        {
            Log("TestDocPlugin DeleteDocument");
            EnsurePathsExist();

            CRM.ReturnInfo res = new CRM.ReturnInfo() { Success = false };
            string path = GetPath(documentInfo);
            if (File.Exists(path))
            {
                if (!CanLock || CanLock && GetCheckoutInfo(path) == "")
                {
                    File.Delete(path);
                    res.Success = true;
                }
            }
            return res;
        }

        /// <summary>
        /// Delete a document template, all language variations and all metadata from the repository
        /// </summary>
        /// <param name="documentTemplateInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type. ("None", "Message", "SoProtocol", "CustomGUI", "Other")</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
		public CRM.ReturnInfo DeleteDocumentTemplate(CRM.IDocumentTemplateInfo documentTemplateInfo, string[] allowedReturnTypes)
        {
            Log("TestDocPlugin DeleteDocumentTemplate");
            CRM.ReturnInfo res = new CRM.ReturnInfo() { Success = false };
            string path = GetPath(documentTemplateInfo);
            if (File.Exists(path))
            {
                File.Delete(path);
                res.Success = true;
            }
            return res;
        }

        /// <summary>
        /// Delete a specific language variation from the document template
        /// </summary>
        /// <param name="documentTemplateInfo">Fully populated document template metadata used to identity the template.</param>
        /// <param name="languageCode">The language variation to delete</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types.</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
		public CRM.ReturnInfo DeleteDocumentTemplateLanguage(CRM.IDocumentTemplateInfo documentTemplateInfo, string languageCode, string[] allowedReturnTypes)
        {
            Log("TestDocPlugin DeleteDocumentTemplateLanguage");
            CRM.ReturnInfo res = new CRM.ReturnInfo() { Success = false };
            string path = GetPath(documentTemplateInfo, languageCode);
            if (File.Exists(path))
            {
                File.Delete(path);
                res.Success = true;
            }
            return res;
        }

        /// <summary>
        /// Execute a custom command on a specified document and version
        /// </summary>
        /// <remarks>
        /// This command is called when the user chooses an action item from a dropdown/context menu. 
        /// It is also reflected in the DocumentAgent service interface, so that custom GUI’s and external 
        /// code can directly execute document plugin commands; this is useful if a plugin also has some 
        /// corresponding custom GUI that needs to execute commands depending on user interaction.
        /// <para/>
        /// The parameter <see param="allowedReturnTypes"/> can be used by clients to hint to the plugin
        /// what kind of return value processing is available. For instance, a mobile client might
        /// only offer None and Message, and this information can be used by the document plugin to adapt
        /// the processing of a command, if it wants to (for instance, use default values instead of
        /// triggering some more advanced workflow).
        /// </remarks>
        /// <param name="documentInfo">Information about the document used from the SuperOffice database</param>
        /// <param name="versionId">Version identifier, blank implies 'latest' version</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type.</param>
        /// <param name="command">Command name, taken from an earlier call to <see cref="GetDocumentCommands"/>
        /// - or any other command name that is understood by the provider. 'Private' commands that
        /// are not declared in GetDocumentCommands but are known to the authors of custom GUI code
        /// or OK.</param>
        /// <param name="additionalData">Array of strings containing whatever additional data the command
        /// may need. This parameter is intended for authors of more complex custom GUI's and works as
        /// a tunnel between the ultimate client and the document plugin. Standard GUI made by SuperOffice,
        /// such as a context menu connected to a document item in an archive, will not populate this
        /// member.<br/>It is strongly suggested that the convention of using name=value for each string
        /// array element be followed here.</param>
        /// <returns>Return value object, specifying failure or success plus any optional, additional processing to be triggered</returns>
        public CRM.ReturnInfo ExecuteDocumentCommand(CRM.IDocumentInfo documentInfo, string versionId, string[] allowedReturnTypes, string command, params string[] additionalData)
        {
            EnsurePathsExist();
            CRM.ReturnInfo res = null;
            if (CanCommands)
            {
                switch (command)
                {
                    case "P":
                        res = new CRM.ReturnInfo() { Success = true, Type = CRM.ReturnType.URL, Value = "https://www.google.com/search?q={0}&source=lnms&tbm=isch&sa=X".FormatWith(documentInfo.Header) };
                        break;
                    case "T":
                        res = new CRM.ReturnInfo() { Success = true, Type = CRM.ReturnType.Message, Value = "You ran Text! command on '{0}'.".FormatWith(documentInfo.Header) };
                        break;
                    case "S":
                        res = new CRM.ReturnInfo() { Success = true, Type = CRM.ReturnType.SoProtocol, Value = "superoffice:contact.main?contact_id={0}".FormatWith(documentInfo.ContactId) };
                        break;
                    default:
                        res = new CRM.ReturnInfo() { Success = false };
                        break;
                }
            }
            return res;
        }

        /// <summary>
        /// Determine if the document exists in the repository
        /// </summary>
        /// <remarks>
        /// The plugin should declare, through the <see cref="SuperOffice.CRM.Documents.Constants.Capabilities.FastExists"/> property,
        /// whether this call is highly efficient or not. If it is efficient, then document archive providers and similar code
        /// will call it when populating an archive, otherwise not.
        /// </remarks>
        /// <param name="documentInfo">Information about the Document</param>
        /// <returns>true if the document exists in the repository, otherwise false</returns>
        public bool Exists(CRM.IDocumentInfo documentInfo)
        {
            Log("TestDocPlugin Exists");
            string path = GetPath(documentInfo);
            bool res = File.Exists(path);
            return res;
        }

        /// <summary>
        /// Get the checkout state of a document
        /// </summary>
        /// <remarks>
        /// This API is called from inside document archive providers if the plugin has declared that it
        /// supports fast fetching of this attribute. If the document plugin does not support locking or
        /// versioning, then the return value should have state NotCheckedOut, associate id 0 and blank name.
        /// </remarks>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <returns>Object that describes the checkout state of the document</returns>
        public CRM.Documents.CheckoutInfo GetCheckoutState(CRM.IDocumentInfo documentInfo)
        {
            Log("TestDocPlugin GetCheckoutState");
            var res = new CheckoutInfo() { State = CheckoutState.LockingNotSupported };
            string path = GetPath(documentInfo);
            if (CanLock)
            {
                res.State = CheckoutState.NotCheckedOut;
                string currentUser = SoContext.CurrentPrincipal.Associate;
                string locker = GetCheckoutInfo(path);
                res.Name = locker;
                if (locker == "")
                    res.State = CheckoutState.NotCheckedOut;
                else
                if (locker == currentUser)
                    res.State = CheckoutState.CheckedOutOwn;
                else
                    res.State = CheckoutState.CheckedOutOther;
            }
            return res;
        }

        /// <summary>
        /// Get a list of custom commands, applicable to a specific document. Note that commands related to
        /// standard locking and versioning operations have their own API calls and are not 'custom commands' in this sense.
        /// </summary>
        /// <remarks>
        /// This API is called before a menu, task button or other GUI item that gives access to document-specific commands is shown.
        /// It is used to populate the GUI with available commands for a particular document, the results are not cached by the GUI.
        /// <para/>
        /// Depending on the return type indicated in the command, the command might be filtered by GUI. More information can
        /// be found in the <see cref="CommandInfo"/> topic.
        /// </remarks>
        /// <param name="documentInfo">Information about the Document</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type.</param>
        /// <returns>Array of command descriptions. If there are no custom commands available, an empty array should be returned.</returns>
        public CommandInfo[] GetDocumentCommands(CRM.IDocumentInfo documentInfo, string[] allowedReturnTypes)
        {
            Log("TestDocPlugin GetDocumentCommands");
            CommandInfo[] res = null;
            if (CanCommands)
            {
                res = new CommandInfo[3];
                res[0] = new CommandInfo() { DisplayName = "Picture!", Name = "P", DisplayTooltip = "Convert into picture", IconHint = "hint", ReturnType = CRM.ReturnType.URL };
                res[1] = new CommandInfo() { DisplayName = "Text!", Name = "T", DisplayTooltip = "Convert into text", IconHint = "", ReturnType = CRM.ReturnType.Message };
                res[2] = new CommandInfo() { DisplayName = "Show!", Name = "S", DisplayTooltip = "Show Contact", IconHint = null, ReturnType = CRM.ReturnType.SoProtocol };
            }
            return res;
        }

        /// <summary>
        /// Map file path to a document id.
        /// </summary>
        /// <param name="documentPathAndName">"C:\SO_ARC\USER\2014.1\foobar.docx"</param>
        /// <returns>123</returns>
        public int GetDocumentIdFromPath(string documentPathAndName)
        {
            return 0;
        }

        /// <summary>
        /// Get the values of certain properties, for a given document
        /// </summary>
        /// <remarks>
        /// Each document can have a number of properties associated with it. A set of standard properties
        /// is defined in the <see cref="SuperOffice.CRM.Documents.Constants.Properties"/> class. Ideally, retrieving properties should
        /// be a lightweight operation.
        /// <para/>
        /// Note that 'properties' are a one-way mechanism where the document plugin provides information about
        /// the document or certain aspects of it. This is not the same as document-specific
        /// metadata, which is handled by the <see cref="LoadMetaData"/> and <see cref="SaveMetaData"/>
        /// methods.
        /// </remarks>
        /// <param name="documentInfo">Information about the Document used from the SuperOffice database</param>
        /// <param name="requestedProperties">Array of property strings, for which values are requested</param>
        /// <returns>Array of name=value pairs, where the name is one of the requested property strings, and the value
        /// is the value of that property for the given document.</returns>
        public Dictionary<string, string> GetDocumentProperties(CRM.IDocumentInfo documentInfo, string[] requestedProperties)
        {
            Log("TestDocPlugin GetDocumentProperties");
            string path = GetPath(documentInfo);
            var props = new Dictionary<string, string>();
            foreach (string prop in requestedProperties)
                props[prop] = "";
            props[Constants.Properties.HasLocking] = CanLock.ToString();
            props[Constants.Properties.HasVersioning] = CanVersion.ToString();
            if (File.Exists(path))
                props[Constants.Properties.LastModified] = SuperOffice.CRM.Globalization.CultureDataFormatter.Encode(File.GetLastWriteTime(path));
            props[Constants.Properties.PreferredOpen] = Constants.Values.Stream;
            return props;
        }

        /// <summary>
        /// Get the list of languages supported by the given template, not including the default (blank) language.
        /// </summary>
        /// <remarks>Used when populating the dropdown list in the admin client or the document dialog.</remarks>
        /// <param name="documentTemplateInfo">The template we are curious about</param>
        /// <returns>Array of ISO codes: ("en-US", "nb-NO", "fr")</returns>
        public string[] GetDocumentTemplateLanguages(CRM.IDocumentTemplateInfo documentTemplateInfo)
        {
            Log("TestDocPlugin GetDocumentTemplateLanguages");
            string path = GetPath(documentTemplateInfo);
            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            var files = Directory.GetFiles(dir, name + " *");
            List<string> langs = new List<string>();
            foreach (var file in files)
            {
                string n = Path.GetFileNameWithoutExtension(file);
                n = n.Replace(name, "");
                n = n.Trim();
                langs.Add(n);
            }

            return langs.ToArray();
        }

        /// <summary>
        /// Get the values of certain properties, for a given document template
        /// </summary>
        /// <remarks>
        /// Each document can have a number of properties associated with it. A set of standard properties
        /// is defined in the <see cref="SuperOffice.CRM.Documents.Constants.Properties"/> class. Ideally, retrieving properties should
        /// be a lightweight operation.
        /// <para/>
        /// Note that 'properties' are a one-way mechanism where the document plugin provides information about
        /// the document or certain aspects of it. This is not the same as document-specific
        /// metadata, which is handled by the <see cref="LoadMetaData"/> and <see cref="SaveMetaData"/>
        /// methods.
        /// </remarks>        
        /// <param name="documentTemplateInfo">Document template record from the SuperOffice database</param>
        /// <param name="requestedProperties">Array of property strings, for which values are requested</param>
        /// <returns>Dictionary of name=value pairs, where the name is one of the requested property strings, and the value
        /// is the value of that property for the given document.</returns>
        public Dictionary<string, string> GetDocumentTemplateProperties(CRM.IDocumentTemplateInfo documentTemplateInfo, string[] requestedProperties)
        {
            Log("TestDocPlugin GetDocumentTemplateProperties");
            var props = new Dictionary<string, string>();
            foreach (string prop in requestedProperties)
                props[prop] = "";
            return props;
        }

        /// <summary>
        /// Get a URL referring to the given document template:  "file:////fileserver/soarc/template/file.ext"
        /// </summary>
        /// <remarks>
        /// Document plugins may support document access via URLs. This call is used to retrieve a url that 
        /// will give the specified access to the document. This URL will be passed to the ultimate client 
        /// (most probably a browser, but could be a text editor application), and control will not return to NetServer.
        /// <para/>
        /// The string returned here should be a fully resolved URL that can be given directly to the editor application.
        /// </remarks>
        /// <param name="documentTemplateInfo">The document template info from database</param>        
        /// <param name="writeableUrl">If true, then the request URL should allow the document editor to write content
        /// back to the repository; otherwise, a url that does not support writeback should be supplied
        /// if possible.</param>
        /// <param name="languageCode">Language variation on the template. May be ignored by the plugin, or used to keep language specific versions of the template.</param>
        /// <returns>URL that gives access to the template: "file:////fileserver/soarc/template/file.ext"</returns>
        public string GetDocumentTemplateUrl(CRM.IDocumentTemplateInfo documentTemplateInfo, bool writeableUrl, string languageCode)
        {
            Log("TestDocPlugin GetDocumentTemplateUrl");
            string path = GetPath(documentTemplateInfo, languageCode);
            Uri uri = new Uri(path);
            return uri.AbsoluteUri; // file:////fileserver/share/dir/file.ext

        }

        /// <summary>
        /// Get a WebDAV-compliant URL referring to the given document
        /// </summary>
        /// <remarks>
        /// Document plugins may support document access via WebDAV. This call is used to retrieve a WebDAV url that 
        /// will give the specified access to the document. This URL will be passed to the ultimate client 
        /// (most probably a text editor application), and control will not return to NetServer.
        /// <para/>
        /// The string returned here should be a fully resolved URL that can be given directly to the editor application.
        /// </remarks>
        /// <param name="incomingInfo">Fully populated document metadata, used to identify the document.</param>
        /// <param name="versionId">Optional version identifier, blank implies 'latest' version</param>
        /// <param name="writeableUrl">If true, then the request URL should allow the document editor to write content
        /// back to the repository; otherwise, a url that does not support writeback should be supplied
        /// if possible.</param>
        /// <returns>" file:////fileserver/share/dir/file.ext"</returns>
        public string GetDocumentUrl(CRM.IDocumentInfo incomingInfo, string versionId, bool writeableUrl)
        {
            Log("TestDocPlugin GetDocumentUrl");
            EnsurePathsExist();
            string path = GetPath(incomingInfo, versionId);
            Uri uri = new Uri(path);
            return uri.AbsoluteUri; // file:////fileserver/share/dir/file.ext
        }

        /// <summary>
        /// Return the length of the physical document. This should be an efficient method
        /// </summary>
        /// <param name="documentInfo">Information about the Document used from the SuperOffice database</param>
        /// <param name="versionId">Version identifier, blank implies 'latest' version</param>
        /// <returns>Physical document length in bytes - this should be the same as the length of the stream
        /// returned by the LoadDocumentStream method.</returns>
		public long GetLength(CRM.IDocumentInfo documentInfo, string versionId)
        {
            string path = GetPath(documentInfo, versionId);
            if (!File.Exists(path))
                return -1;
            var fi = new FileInfo(path);
            long res = fi.Length;
            return res;
        }

        /// <summary>
        /// Get a list of capabilities (functionality) supported by this document plugin
        /// </summary>
        /// <remarks>
        /// The purpose of this call is to enable NetServer and clients to determine what functionality this plugin can offer. 
        /// Plugins should populate the return array with all capabilities they know about. NetServer will call this API only once.
        /// <para/>
        /// As an example of use, the Document archive provider inside NetServer will look at plugin capabilities, 
        /// and read document properties as appropriate. 
        /// <para/>
        /// i.e. if “fast-lock-status=false”, then the archive provider 
        /// will not call the IsCheckedOut(externalReference) function. Otherwise it will make the call (if the client has requested
        /// the appropriate column in the GUI), so that the user can see which documents are checked out.
        /// <para/>
        /// String constants for capabilities are available in the <see cref="SuperOffice.CRM.Documents.Constants.Capabilities"/> static class.
        /// </remarks>
        /// <returns>Dictionary of name=value strings listing all known capabilities and their values</returns>
        public Dictionary<string, string> GetPluginCapabilities()
        {
            Log("TestDocPlugin GetPluginCapabilites");
            EnsurePathsExist();

            var capabilities = new Dictionary<string, string>();
            capabilities[Constants.Capabilities.CanCreateDocumentTemplates] = CanCreateDocumentTemplates.ToString();
            capabilities[Constants.Capabilities.Versioning] = CanVersion.ToString();
            capabilities[Constants.Capabilities.Locking] = CanLock.ToString();
            capabilities[Constants.Capabilities.FastVersionList] = CanVersion.ToString();
            capabilities[Constants.Capabilities.FastLockStatus] = CanLock.ToString();
            capabilities[Constants.Capabilities.FastExists] = true.ToString();

            return capabilities;
        }

        /// <summary>
        /// Get a list of supported document template types. 
        /// </summary>
        /// <returns>An dictionary of key=display-name for supported document types for template. Empty dictionary if no document types supported.</returns>
        public Dictionary<int, string> GetSupportedDocumentTypesForDocumentTemplates()
        {
            var res = new Dictionary<int, string>();
            res[1] = "Word Document";
            res[2] = "Text document";
            res[3] = "Picture";
            return res;
        }

        /// <summary>
        /// Get the "extension" for the template, i.e., what the file extension would have been - to 
        /// help identify the stream content format
        /// </summary>
        /// <remarks>
        /// Template documents are generally created in text editors and stored as files of some kind. The
        /// file extension indicates the kind of document - doc, docx, xls, txt, and so on. While the template
        /// may be stored inside the document repository as any kind of data byte collection, a concept
        /// akin to the file extension is still needed to help identify the document format, ahead of actually
        /// reading the template content.
        /// </remarks>
        /// <param name="documentTemplateInfo">Document template info: contains the extref/filename, template id, mime type.</param>    
        /// <returns>String equivalent to a file extension, for instance ".docx"</returns>
        public string GetTemplateExtension(CRM.IDocumentTemplateInfo documentTemplateInfo)
        {
            Log("TestDocPlugin GetTemplateExtension");
            string path = GetPath(documentTemplateInfo);
            string ext = Path.GetExtension(path);
            return ext;
        }

        /// <summary>
        /// Get the list of current versions for the given document
        /// </summary>
        /// <remarks>
        /// The list should not include an “in-work” version, if the document is currently checked out – only 
        /// versions visible and accessible to any authorized user.
        /// <para/>
        /// If the document plugin does not support versioning, then this call should return an empty array.
        /// </remarks>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <returns>Array of objects describing the existing, committed versions for this document</returns>
        public CRM.Documents.VersionInfo[] GetVersionList(CRM.IDocumentInfo documentInfo)
        {
            Log("TestDocPlugin GetVersionList");
            var res = new List<VersionInfo>();
            if (CanVersion)
            {
                string path = GetPath(documentInfo, "*");
                string dir = Path.GetDirectoryName(path);
                string wild = Path.GetFileName(path);
                string ext = Path.GetExtension(path);
                string prefix = wild.Substring(0, wild.IndexOf("*"));
                var files = Directory.GetFiles(dir, wild);
                foreach (var file in files)
                {
                    string name = file.Replace(dir, "");
                    string ver = name.Replace(prefix, "");
                    ver = ver.Replace("\\", "");
                    ver = ver.Replace(ext, "");
                    DateTime lastMod = File.GetLastWriteTime(file);
                    var verInfo = new VersionInfo() { DisplayText = "Ver " + ver + " " + lastMod.ToShortDateString(), DocumentId = documentInfo.DocumentId, ExternalReference = documentInfo.ExternalReference, VersionId = ver, CheckedInDate = lastMod };
                    res.Add(verInfo);
                }
            }
            return res.ToArray();
        }

        /// <summary>
        /// Get document content as a stream. NetServer will read-to-end and close this stream.
        /// </summary>
        /// <remarks>
        /// It is up to the document plugin whether it can open a stream directly into the underlying repository, 
        /// or whether it has to extract the document to some temporary area and then stream that – 
        /// however, the fewer buffers the better.
        /// </remarks>
        /// <param name="incomingInfo">Fully populated document metadata, used to identify the document.</param>
        /// <param name="versionId">Optional version identifier, blank implies 'latest' version</param>
        /// <returns>Document content stream</returns>
        public System.IO.Stream LoadDocumentStream(CRM.IDocumentInfo incomingInfo, string versionId)
        {
            Log("TestDocPlugin LoadDocumentStream");
            EnsurePathsExist();
            string path = GetPath(incomingInfo, versionId);
            return File.OpenRead(path);
        }

        /// <summary>
        /// Get the document template content as a stream. NetServer will read-to-end and close this stream
        /// </summary>
        /// <remarks>
        /// Document templates may be stored in a repository, with or without special content tags.
        /// Because a document template does not have a corresponding document record within
        /// SuperOffice, there is no documentId to identify it.
        /// <para/>
        /// This call is used by NetServer to retrieve a document template based on either
        /// an externalreference value stored in the corresponding DocTmpl.Filename field,
        /// or the Id of the doctmpl record itself. The document plugin is free
        /// to use either method of identification.
        /// <para/>
        /// Mail templates are passed in using extref = "filename=xyz&amp;allowPersonal=1" and docTemplateId = 0
        /// </remarks>
        /// <param name="documentTemplateInfo">Document template info: contains the extref/filename, template id, mime type.
        /// TemplateInfo.Id = 0 when archiving mail messages. 
        /// </param>
        /// <param name="languageCode">Language (en-US, nb-NO, etc) that the user is using in the user interface. Can be used to select language-specific templates.</param>
        /// <returns>Stream containing the template content. Null if no suitable template found.</returns>
        public System.IO.Stream LoadDocumentTemplateStream(CRM.IDocumentTemplateInfo documentTemplateInfo, string languageCode)
        {
            Log("TestDocPlugin LoadDocumentTemplateStream");
            EnsurePathsExist();
            string path = GetPath(documentTemplateInfo, languageCode);
            // fall back to neutral culture if language specific template not found
            if (!File.Exists(path))
                path = GetPath(documentTemplateInfo, "");
            return File.OpenRead(path);
        }

        /// <summary>
        /// Retrieve metadata owned by the plugin/repository, related to one document (excluding version-dependent metadata)
        /// </summary>
        /// <remarks>
        /// A document plugin may consume and provide an arbitrary number of metadata fields. These are placed
        /// in a string, string dictionary, representing name/value pairs. This call should <b>not</b>
        /// retrieve metadata related to any particular version, only data related to the document as a whole.
        /// <para/>
        /// The NetServer service call DocumentAgent.GetDocumentEntity will (among other things) result in a call 
        /// to this API to populate the ExtraFields property of the document entity carrier. Note, however, 
        /// that the carrier handed over to the client may contain other fields in addition to those 
        /// supplied by the document plugin, since the extrafields mechanism is generic and 
        /// there may be other metadata providers along the line.
        /// <para/>
        /// Attribute names should be prefixed with the name of the document plugin, to maintain global uniqueness.
        /// <para/>
        /// It is <b>strongly suggested</b> that non-string data be formatted according to the rules followed by the 
        /// <see cref="SuperOffice.CRM.Globalization.CultureDataFormatter"/> class, to avoid problems
        /// when parsing dates and floating-point types between different cultures and platforms.
        /// </remarks>
        /// <param name="documentInfo">Document info used by the document plugin</param>
        /// <returns>Dictionary of name=value strings, each representing one key and one value.
        /// Always throws InvalidOperationException.
        /// </returns>
        public Dictionary<string, string> LoadMetaData(CRM.IDocumentInfo documentInfo)
        {
            return null;
        }

        /// <summary>
        /// Retrieve metadata owned by the plugin/repository, related to one particular
        /// version of one document
        /// </summary>
        /// <remarks>
        /// A document plugin may consume and provide an arbitrary number of metadata fields. These are placed
        /// in a string, string dictionary, representing name/value pairs. This call should only retrieve
        /// metadata related to a version.
        /// <para/>
        /// The NetServer service call DocumentAgent.GetDocumentEntity will (among other things) result in a call 
        /// to this API to populate the VersionInfo property of the document entity carrier. 
        /// <para/>
        /// Attribute names should be prefixed with the name of the document plugin, to maintain global uniqueness.
        /// <para/>
        /// It is <b>strongly suggested</b> that non-string data be formatted according to the rules followed by the 
        /// <see cref="SuperOffice.CRM.Globalization.CultureDataFormatter"/> class, to avoid problems
        /// when parsing dates and floating-point types between different cultures and platforms.
        /// <para/>
        /// To efficiently retrieve information about <b>all</b> versions, use the 
        /// <see cref="SuperOffice.CRM.Documents.IDocumentPlugin2.GetVersionList"/> method, instead of iterating
        /// over this method.
        /// </remarks>
        /// <param name="documentInfo">Information about the document</param>
        /// <param name="versionId">Version identifier, blank implies 'latest' version</param>
        /// <returns>Fully populated version info structure.
        /// Always throws InvalidOperationException.
        /// </returns>
        public CRM.Documents.VersionInfo LoadVersionInfo(CRM.IDocumentInfo documentInfo, string versionId)
        {
            if (!CanVersion)
                return null;
            string path = GetPath(documentInfo, versionId);
            CRM.Documents.VersionInfo res = null;
            if (File.Exists(path))
            {
                var lastMod = File.GetLastWriteTime(path);
                res = new CRM.Documents.VersionInfo() { VersionId = versionId, ExternalReference = path, DocumentId = documentInfo.DocumentId, CheckedInDate = lastMod, DisplayText = "Ver " + versionId + " " + lastMod.ToShortDateString() };
            }
            return res;
        }

        /// <summary>
        /// Rename a document in the repository
        /// </summary>
        /// <remarks>
        /// The document name should be changed from the existing to the new name. However,
        /// if the new name is not valid (or collides with an existing name of some other
        /// document), then the plugin should amend the name to a valid one and return
        /// it, instead of throwing an exception.
        /// </remarks>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <param name="suggestedNewName">Suggested new document name</param>
        /// <returns>Actual new document name, limited to 254 characters
        /// If renaming was not occured this will return an empty string</returns>
		public string RenameDocument(CRM.IDocumentInfo documentInfo, string suggestedNewName)
        {
            Log("TestDocPlugin RenameDocument - not impl");
            throw new NotImplementedException();
        }

        /// <summary>
        /// Save the stream as the document content in the repository; depending on the state, this
        /// may imply creating a temporary save pending a final checkin, or an immediately visible result.
        /// </summary>
        /// <remarks>
        /// If the document is currently checked out to the current user, then the stream should be saved, 
        /// but this call does not imply the automatic creation of a new version (visible to other users) 
        /// or automatic checkin. However, it is an advantage if subsequent GetDocument calls made by 
        /// the same user using the same key return the latest known content – while other users see 
        /// the latest checked-in version.
        /// <para/>
        /// If the plugin does not support locking agnd versioning (or such semantics are not requested, see below), 
        /// then each call to this API overwrites 
        /// any prior content completely and becomes the new, official content immediately. The Save operation 
        /// should be atomic, and should not destroy earlier content if it fails.
        /// <para/>
        /// If locking is supported and requested, the document is checked out and some other associate than the one 
        /// that has checked it out calls this API, a failure message should be returned.
        /// </remarks>
        /// <param name="incomingInfo">Incoming document metadata, used to identify the document. Metadata
        /// changes are <b>not</b> to be checked or saved by this operation - only the document stream is saved.</param>
        /// <param name="content">Document content, a binary stream about which nothing is assumed. The
        /// document plugin should read-to-end and close this stream.</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type.</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
		public CRM.ReturnInfo SaveDocumentFromStream(CRM.IDocumentInfo incomingInfo, string[] allowedReturnTypes, System.IO.Stream content)
        {
            Log("TestDocPlugin SaveDocumentFromStream");
            EnsurePathsExist();
            CRM.ReturnInfo res = new CRM.ReturnInfo() { Success = false };
            string path = GetPath(incomingInfo);
            // implicit checkout of file
            string currentUser = SoContext.CurrentPrincipal.Associate;
            if (CanLock && GetCheckoutInfo(path) == "")
                SetCheckoutInfo(path, currentUser);
            if (!CanLock || CanLock && GetCheckoutInfo(path) == currentUser)
            {
                if (File.Exists(path))
                    File.Delete(path);
                var file = File.Create(path);
                content.CopyTo(file);
                content.Close();
                file.Close();
                res.Success = true;
                res.ExternalReference = incomingInfo.ExternalReference;
            }
            else
            {
                res.Type = CRM.ReturnType.Message;
                res.Value = "Locked by " + GetCheckoutInfo(path);
            }
            return res;
        }

        /// <summary>
        /// Create or update the document template contents. Usually used when uploading a file to a new document template.
        /// </summary>
        /// <param name="templateInfo">Name and tooltip from the document template record in the database. The ExtRef/Filename may be set if this is an edit rather than an add.</param>
        /// <param name="content">Stream containing file content</param>
        /// <param name="languageCode">Language variation on the template. May be ignored by the plugin, or used to keep language specific versions of the template.</param>
        /// <returns>Template information with ExtRef/Filename and MimeType filled in. These values are saved in the DocTmpl record.</returns>
        public CRM.Documents.TemplateInfo SaveDocumentTemplateStream(CRM.IDocumentTemplateInfo templateInfo, System.IO.Stream content, string languageCode)
        {
            Log("TestDocPlugin SaveDocumentTemplateStream");
            EnsurePathsExist();
            TemplateInfo res = new TemplateInfo();
            string path = GetPath(templateInfo, languageCode);
            // implicit checkout of file

            var file = File.OpenWrite(path);
            content.CopyTo(file);
            content.Close();
            file.Close();

            res.Name = templateInfo.Name;
            res.PluginId = 123;
            res.Description = templateInfo.Tooltip;
            res.ExternalReference = path;

            return res;
        }

        /// <summary>
        /// Store/update plugin-dependent document metadata in the repository
        /// </summary>
        /// <remarks>
        /// This call is made when the document metadata should be stored, and is the complement of the
        /// <see cref="SuperOffice.CRM.Documents.IDocumentPlugin2.LoadMetaData"/> method.
        /// The document plugin should extract whatever elements it 
        /// recognizes from the pluginData name/value dictionary. Failure to recognize an element should not cause an exception, 
        /// as there may be other plugins along the line (not document plugins, but service-level field providers) that own the data. 
        /// Likewise, absence of a value should be taken to imply “no change” to that value - not "delete".
        /// <para/>
        /// It is <b>strongly suggested</b> that non-string data be formatted according to the rules followed by the 
        /// <see cref="SuperOffice.CRM.Globalization.CultureDataFormatter"/> class, to avoid problems
        /// when parsing dates and floating-point types between different cultures and platforms.
        /// 
        /// SoArc plugin does not use metadata.
        /// </remarks>
        /// <param name="incomingInfo">SuperOffice document information. Note that the plugin is <b>not</b> responsible
        /// for storing this data; however, it is allowed to look at it, in case it influences how the document
        /// is stored. However, it should always be possible to retrieve a document using the ExternalReference
        /// or DocumentId keys alone.</param>
        /// <param name="pluginData">Name/value dictionary containing metadata</param>
		public void SaveMetaData(CRM.IDocumentInfo incomingInfo, Dictionary<string, string> pluginData)
        {
            // yawn
        }

        /// <summary>
        /// Store/update plugin-dependent document version metadata in the repository
        /// </summary>
        /// <remarks>
        /// This call is made when the document <b>version</b> metadata should be stored, and is the complement of the
        /// <see cref="SuperOffice.CRM.Documents.IDocumentPlugin2.LoadVersionInfo"/> method.
        /// The document plugin should extract whatever elements it 
        /// recognizes from the pluginData name/value dictionary. Failure to recognize an element should not cause an exception, 
        /// as there may be other plugins along the line (not document plugins, but service-level field providers) that own the data. 
        /// Likewise, absence of a value should be taken to imply “no change” to that value - not "delete".
        /// <para/>
        /// It is <b>strongly suggested</b> that non-string data be formatted according to the rules followed by the 
        /// <see cref="SuperOffice.CRM.Globalization.CultureDataFormatter"/> class, to avoid problems
        /// when parsing dates and floating-point types between different cultures and platforms.
        /// </remarks>
        /// <param name="documentInfo">Document that version is being saved on</param>
        /// <param name="versionInfo">Version information to be saved</param>
        public void SaveVersionInfo(CRM.IDocumentInfo documentInfo, CRM.Documents.VersionInfo versionInfo)
        {
            // not really gonna happen
        }

        /// <summary>
        /// Undo (abandon) a checkout
        /// </summary>
        /// <remarks>
        /// If the document plugin supports locking and the requesting user is the one who checked out the document, 
        /// then any content saved since the checkout should be discarded and the checkout state reset. 
        /// The content will be as before checkout. 
        /// <para/>
        /// Calls by other users should result in failure and no state change – except if the calling user has the right to force an undo
        /// <para/>
        /// If the document plugin does not support locking or versioning, then this call should perform no action.
        /// </remarks>
        /// <param name="documentInfo">Fully populated document metadata, used to identify the document. Usefully contains ExternalReference and Filename properties.</param>
        /// <param name="allowedReturnTypes">Array of names of allowed return types; if this array is
        /// empty then no limits are placed on return type.</param>
        /// <returns>Return value, indicating success/failure and any optional processing to be performed</returns>
        public CRM.ReturnInfo UndoCheckoutDocument(CRM.IDocumentInfo documentInfo, string[] allowedReturnTypes)
        {
            Log("TestDocPlugin UndoCheckoutDocument");
            var res = new CRM.ReturnInfo() { Success = false };
            if (CanLock)
            {
                string currentUser = SoContext.CurrentPrincipal.Associate;
                string path = GetPath(documentInfo);
                string locker = GetCheckoutInfo(path);
                if (locker == "")
                {
                    res.Success = false;
                    res.Type = CRM.ReturnType.Message;
                    res.Value = "Not checked out";
                }
                else
                    if (locker == currentUser)
                    res.Success = true;
                else
                {
                    res.Success = false;
                    res.Type = CRM.ReturnType.Message;
                    res.Value = "Checked out to " + locker;
                }

                if (res.Success)
                {
                    // find last available ver number
                    int verNum = 1;
                    string verPath = GetPath(documentInfo, verNum.ToString());
                    string lastPath = verPath;
                    while (File.Exists(verPath))
                    {
                        lastPath = verPath;
                        verNum++;
                        verPath = GetPath(documentInfo, verNum.ToString());
                    }
                    // Revert the current to the last check-in
                    File.Copy(lastPath, path);

                    SetCheckoutInfo(path, "");
                }
            }
            return res;
        }


        #region private helpers

        private void Log(string msg)
        {
#if DEBUG
            msg = msg.Replace("\\", "/");
            System.Diagnostics.Debug.WriteLine(msg);
            SuperOffice.Diagnostics.SoLogger.LogInformation(GetType(), msg, "", true);
#endif
        }

        private string GetPath(CRM.IDocumentInfo doc)
        {
            string name = doc.ExternalReference;
            if (string.IsNullOrEmpty(name))
                name = "blank";
            string path = Path.Combine(RootPath, "Documents", name);
            return path;
        }

        private string GetPath(CRM.IDocumentInfo doc, string version)
        {
            if (string.IsNullOrEmpty(version))
                return GetPath(doc);

            string name = Path.GetFileNameWithoutExtension(doc.ExternalReference);
            string ext = Path.GetExtension(doc.ExternalReference);
            string path = Path.Combine(RootPath, "Documents", name + " v" + version + ext);
            return path;
        }

        private string GetPath(CRM.IDocumentTemplateInfo doc)
        {
            string path = Path.Combine(RootPath, "Templates", doc.ExternalReference);
            return path;
        }

        private string GetPath(CRM.IDocumentTemplateInfo doc, string langCode)
        {
            string name = Path.GetFileNameWithoutExtension(doc.ExternalReference);
            string ext = Path.GetExtension(doc.ExternalReference);
            string path = Path.Combine(RootPath, "Templates", name + langCode + ext);
            return path;
        }

        private void EnsurePathsExist()
        {
            string path = RootPath;
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            path = Path.Combine(RootPath, "Templates");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            path = Path.Combine(RootPath, "Documents");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
        }

        private bool LockedByOtherProcess(string path )
		{
			// check for file locks
			try
			{
				if (File.Exists(path))
				{
					// See if Word or something is holding the file open
					using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
					{
						fs.Close();
					}
				}
			}
			catch (Exception)
			{
				return true;
			}
			return false;
		}
		private string GetCheckoutInfo(string path)
		{
			string lockPath = path + ".lock";
			if (LockedByOtherProcess(path))
				return "Other process";
			if (!File.Exists(lockPath) )
				return "";
			string locker = File.ReadAllText(lockPath);
			return locker;
		}

		private void SetCheckoutInfo(string path, string locker)
		{
			string lockPath = path + ".lock";
			if (string.IsNullOrEmpty(locker))
				File.Delete(lockPath);
			else
				File.WriteAllText(lockPath, locker);
		}
        #endregion
    }
}
