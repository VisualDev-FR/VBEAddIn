[Setup]

; !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
; BEGIN CUSTOMIZATION SECTION: Customize these constants
; !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

; Ensure that you use YOUR OWN APP_ID; DO NOT REUSE THIS ONE
#define APP_ID "{{ead31314-ea00-455a-8ff8-750ec04692e2}"

#define APP_NAME "VBEAddIn 1.0"
#define DEST_SUB_DIR "VBEAddIn"
#define CONNECT_CLASS_FULL_NAME "VBEAddIn.Connect"
#define COMPANY_NAME "TMenanteau"
#define RUNTIME_VERSION "v2.0.50727"
#define COPYRIGHT_YEAR "2022"

#define INTEROP_OFFICE_FILE_NAME "VBEAddIn.Interop.Office11.dll"
#define INTEROP_STDOLE_FILE_NAME "VBEAddIn.Interop.Stdole.dll"
#define INTEROP_EXTENSIBILITY_FILE_NAME "VBEAddIn.Interop.Extensibility.dll"
#define INTEROP_VBA_EXTENSIBILITY_FILE_NAME "VBEAddIn.Interop.VBAExtensibility.dll"
#define DLL_FILE_NAME "VBEAddIn.dll"
#define OUTPUT_FOLDER_NAME ".\Setup"
#define SETUP_FILE_NAME "VBEAddInSetup"
#define VERSION "1.0.0.00"
#define CONNECT_PROGID "VBEAddIn.Connect"
#define CONNECT_CLSID "{875B3991-9A51-48AC-A328-ABE02EB53279}"
#define USERCONTROL_PROGID "VBEAddIn.UserControlHost"
#define USERCONTROL_CLSID "{0607197B-887A-4364-9386-ADAAE7227FD9}"
#define ASSEMBLY_FULL_NAME "MyVBAAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=869cad219d7a35e2"

; !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
; END CUSTOMIZATION SECTION
; !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

ArchitecturesAllowed=x86 x64

; The setup must run in 64-bit mode on x64 systems to allow the installation of the VBA add-in for Office 2010 64-bit
ArchitecturesInstallIn64BitMode=x64 

AppID={#APP_ID}
VersionInfoVersion={#VERSION}
OutputBaseFilename={#SETUP_FILE_NAME}
OutputDir={#OUTPUT_FOLDER_NAME}
PrivilegesRequired=lowest
MinVersion=0,6.0
AppName={#APP_NAME}
AppVerName={#APP_NAME}
DefaultGroupName={#APP_NAME}
AppPublisher={#COMPANY_NAME}
DefaultDirName={localappdata}\{#COMPANY_NAME}\{#DEST_SUB_DIR}
Compression=lzma/Max
SolidCompression=true
DisableReadyPage=true
ShowLanguageDialog=no
UninstallLogMode=append
DisableProgramGroupPage=true
VersionInfoCompany={#COMPANY_NAME}
AppCopyright=Copyright � {#COPYRIGHT_YEAR} {#COMPANY_NAME}
AlwaysUsePersonalGroup=true
InternalCompressLevel=Ultra
AllowNoIcons=true
DisableDirPage=true
LanguageDetectionMethod=locale

[Languages]
; USE ENGLISH AS THE FIRST LANGUAGE!!!
Name: English; MessagesFile: compiler:Default.isl
Name: Spanish; MessagesFile: compiler:Languages\Spanish.isl

[Types]
Name: Custom; Description: Custom; Flags: iscustom

[Files]
Source: bin\release\{#INTEROP_OFFICE_FILE_NAME};            DestDir: {app}; Flags: ignoreversion;
Source: bin\release\{#INTEROP_STDOLE_FILE_NAME};            DestDir: {app}; Flags: ignoreversion;
Source: bin\release\{#INTEROP_VBA_EXTENSIBILITY_FILE_NAME}; DestDir: {app}; Flags: ignoreversion;
Source: bin\release\{#INTEROP_EXTENSIBILITY_FILE_NAME};     DestDir: {app}; Flags: ignoreversion;
Source: bin\release\{#DLL_FILE_NAME};                       DestDir: {app}; Flags: ignoreversion; AfterInstall: RegisterAddin()

[UninstallDelete]
Name: {app}; Type: filesandordirs

[CustomMessages]
English.NETFramework20NotInstalled=Microsoft .NET Framework 2.0 installation was not detected. 
Spanish.NETFramework20NotInstalled=No se encontro la instalacion de Microsoft .NET Framework 2.0. 

[Run]

[UninstallRun]

[Code]
var
   m_IDEsPage: TWizardPage;
   m_IDEsCheckListBox: TNewCheckListBox;

function IsVBA64Installed(): Boolean;
var
   sBitness: String;
begin

   Result := False
   
   if IsWin64() then
      begin
         if RegQueryStringValue(HKLM, 'Software\Microsoft\Office\14.0\Outlook', 'Bitness', sBitness) then
            begin
               if sBitness = 'x64' then
                  begin
                     Result := True
                  end;
            end;
      end;
end;

function IsVBA32Installed(): Boolean;
begin

   if RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\Office\10.0') then
      Result := True
   else if RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\Office\11.0') then
      Result := True
   else if RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\Office\12.0') then
      Result := True
   else if RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\Office\14.0') then
      Result := not IsVBA64Installed()
   else
      Result := False

end;

function IsVBA32Selected(): Boolean;
begin
   Result := m_IDEsCheckListBox.Checked[0]
end;

function IsVBA64Selected(): Boolean;
begin
   Result := m_IDEsCheckListBox.Checked[1]
end;

procedure IDEsCheckListBoxOnClickCheck(Sender: TObject);
var
   index: Integer;
begin

   WizardForm.NextButton.Enabled := False;

   for index := 0 to m_IDEsCheckListBox.Items.Count - 1 do
      begin
         if m_IDEsCheckListBox.Checked[index] then
            WizardForm.NextButton.Enabled := True;
      end; 
end;

//***************************************************************************************************
// InnoSetup event function                                                                        
//***************************************************************************************************
function InitializeSetup(): Boolean;
var
   iErrorCode: Integer;
begin
   // Detect if Microsoft .NET Framework 2.0 is installed
   if Not RegKeyExists(HKLM, 'SOFTWARE\Microsoft\.NETFramework\v2.0.50727') then
      begin
         MsgBox(ExpandConstant('{cm:NETFramework20NotInstalled}'), mbCriticalError, mb_Ok);
         ShellExec('open', 'http://msdn.microsoft.com/en-us/netframework/aa731542', '', '', SW_SHOW, ewNoWait, iErrorCode) 
         Result := False;
      end
   else
      begin
         Result := True;
      end
end;

procedure CreateComponentsPage();
var
   IDEsLabel: TLabel;
   bVBA32Enabled: Boolean;
   bVBA64Enabled: Boolean;
   bVBA32Checked: Boolean;
   bVBA64Checked: Boolean;
begin

   m_IDEsPage := CreateCustomPage(wpSelectComponents, SetupMessage(msgWizardSelectComponents), SetupMessage(msgSelectComponentsDesc));
   
   IDEsLabel := TLabel.Create(m_IDEsPage);
   IDEsLabel.Caption := SetupMessage(msgSelectComponentsLabel2);
   IDEsLabel.Width := m_IDEsPage.SurfaceWidth;
   IDEsLabel.Height := ScaleY(40);
   IDEsLabel.AutoSize := False;
   IDEsLabel.WordWrap := True;
   IDEsLabel.Parent := m_IDEsPage.Surface;

   m_IDEsCheckListBox := TNewCheckListBox.Create(m_IDEsPage);
   m_IDEsCheckListBox.Top := IDEsLabel.Top + IDEsLabel.Height + ScaleY(8);
   m_IDEsCheckListBox.Width := m_IDEsPage.SurfaceWidth;
   m_IDEsCheckListBox.Height := ScaleX(100);
   m_IDEsCheckListBox.Flat := True;
   m_IDEsCheckListBox.Parent := m_IDEsPage.Surface;
   m_IDEsCheckListBox.OnClickCheck := @IDEsCheckListBoxOnClickCheck;

   bVBA32Enabled := IsVBA32Installed();
   bVBA64Enabled := IsVBA64Installed();

   bVBA32Checked := bVBA32Enabled;
   bVBA64Checked := bVBA64Enabled;
  
   m_IDEsCheckListBox.AddCheckBox('My VBA Add-in - VBA Editor 32-bit (Office 2003/2007/2010 32-bit)', '', 0, 
     bVBA32Checked, bVBA32Enabled, False, True, nil)
   m_IDEsCheckListBox.AddCheckBox('My VBA Add-in - VBA Editor 64-bit (Office 2010 64-bit)', '',           0, 
      bVBA64Checked, bVBA64Enabled, False, True, nil)

end;

//***************************************************************************************************
// InnoSetup event function                                                                        
//***************************************************************************************************
procedure InitializeWizard();
begin

   CreateComponentsPage();

end;

procedure RegisterCOMClass(const iRootKey: Integer; const sProgID: String; const sCLSID: String; const sClass: String; 
  const sFileName: String; const sAssemblyFullName: String);
var
   sCodeBase: String;
   sFileFullName: String;
begin

   sFileFullName := ExpandConstant('{app}') + '\' + sFileName;
   StringChangeEx(sFileFullName, '\', '/', True);
   sCodeBase := 'file:///' + sFileFullName;

   RegWriteStringValue(iRootKey, 'Software\Classes\' + sProgID, '', sProgID);
   RegWriteStringValue(iRootKey, 'Software\Classes\' + sProgID + '\CLSID', '', sCLSID);

   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID, '', sClass);

   (* Do not add this implemented category that identifies .NET components:

      RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\Implemented Categories', '', '');
      RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + 
        '\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}', '', '');
   
      because it is not absolutely needed and creating it could cause the following error:

      FIX: "Access to the Registry Key Denied" Error Message When You Register .NET Assembly for COM Interop
      http://support.microsoft.com/kb/327507
   *)

   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\InprocServer32', '', 'mscoree.dll');
   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\InprocServer32', 'Assembly', sAssemblyFullName);
   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\InprocServer32', 'Class', sClass);
   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\InprocServer32', 'CodeBase', sCodeBase);
   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\InprocServer32', 'RuntimeVersion', '{#RUNTIME_VERSION}');
   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\InprocServer32', 'ThreadingModel', 'Both');

   RegWriteStringValue(iRootKey, 'Software\Classes\CLSID\' + sCLSID + '\ProgId', '', sProgID);

end;

procedure UnregisterCOMClass(const iRootKey: Integer; const sProgID: String; const sCLSID: String);
begin

   if RegKeyExists(iRootKey, 'Software\Classes\' + sProgID) then
      begin
         RegDeleteKeyIncludingSubkeys(iRootKey, 'Software\Classes\' + sProgID);
      end;

   if RegKeyExists(iRootKey, 'Software\Classes\CLSID\' + sCLSID) then   
      begin
         RegDeleteKeyIncludingSubkeys(iRootKey, 'Software\Classes\CLSID\' + sCLSID);
      end;

end;

procedure RegisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String);
begin

   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'FriendlyName', '{#APP_NAME}');
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'Description' , '{#APP_NAME}');
   RegWriteDWordValue (iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'LoadBehavior', 0);

end;

procedure UnregisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String);
begin

   if RegKeyExists(iRootKey, sAddinSubKey + '\' + sProgIDConnect) then
      begin
         RegDeleteKeyIncludingSubkeys(iRootKey, sAddinSubKey + '\' + sProgIDConnect);
      end;

end;

procedure RegisterAddinForCOM(const iRootKey: Integer; const sProgIDConnect: String; const sCLSIDConnect: String; 
  const sConnectClassFullName: String; const sFileName: String; const sAssemblyFullName: String);
begin

   RegisterCOMClass(iRootKey, sProgIDConnect, sCLSIDConnect, sConnectClassFullName, sFileName, sAssemblyFullName);

   // Here you would register toolwindow usercontrols if required
     
end;

procedure UnregisterAddinForCOM(const iRootKey: Integer; const sProgIDConnect: String; const sCLSIDConnect: String);
begin

   UnregisterCOMClass(iRootKey, sProgIDConnect, sCLSIDConnect);

   // Here you would unregister toolwindow usercontrols if required

end;

procedure RegisterAddin();
begin
   
   if IsVBA32Selected() then
      begin
         RegisterAddinForCOM(HKCU32, '{#CONNECT_PROGID}', '{#CONNECT_CLSID}', '{#CONNECT_CLASS_FULL_NAME}', 
           '{#DLL_FILE_NAME}', '{#ASSEMBLY_FULL_NAME}');
         RegisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#CONNECT_PROGID}');
      end;

   if IsVBA64Selected() then
      begin
         RegisterAddinForCOM(HKCU64, '{#CONNECT_PROGID}', '{#CONNECT_CLSID}', '{#CONNECT_CLASS_FULL_NAME}', 
            '{#DLL_FILE_NAME}', '{#ASSEMBLY_FULL_NAME}');
         RegisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#CONNECT_PROGID}');
      end;

end;

procedure UnregisterAddin();
begin

   UnregisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#CONNECT_PROGID}');
   UnregisterAddinForCOM(HKCU32, '{#CONNECT_PROGID}', '{#CONNECT_CLSID}');

   if IsWin64() then
      begin
         UnregisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#CONNECT_PROGID}');
         UnregisterAddinForCOM(HKCU64, '{#CONNECT_PROGID}', '{#CONNECT_CLSID}');
      end
end;

//***************************************************************************************************
// InnoSetup event function                                                                        
//***************************************************************************************************
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin

   if CurUninstallStep = usUninstall then
      begin
         UnregisterAddin()
      end;

end;