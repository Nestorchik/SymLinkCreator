unit SymLinkUnit;

interface

uses
  Windows, SysUtils, Controls, Forms, Messages, Dialogs, Grids, Clipbrd, ExtCtrls, Menus, StdCtrls, Classes, Graphics,
  Vcl.ComCtrls, IdBaseComponent, IdMessage, Vcl.TabNotBk, DragDrop, DropSource, DragDropFile, DropTarget, DragDropFormats,
  ImgList, ActnList, Actions, ImageList, Types, example, Vcl.Buttons, Vcl.ToolWin, Vcl.OleCtrls, SHDocVw, Vcl.Imaging.pngimage,
  SymLinkINI, IOUtils, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnPopup, Vcl.WinXCtrls;

type

  TSymLinkForm = class(TForm)
    pMenu: TPopupMenu;
    mSave: TMenuItem;
    mNew: TMenuItem;
    mQuit: TMenuItem;
    SaveFileDialog: TSaveDialog;
    mCopyAll: TMenuItem;
    mCopyNames: TMenuItem;
    MainMenu: TMainMenu;
    mmFile: TMenuItem;
    mmNew: TMenuItem;
    mmOpen: TMenuItem;
    mmSave: TMenuItem;
    mmQuit: TMenuItem;
    mmFileList: TMenuItem;
    mmClear: TMenuItem;
    mmCopyAll: TMenuItem;
    mmCopyOnlyNames: TMenuItem;
    OpenFileDialog: TOpenDialog;
    TabNotebook: TTabbedNotebook;
    mDelete: TMenuItem;
    mmDelClick: TMenuItem;
    msgTimer: TTimer;
    ListBox: TListBox;
    mOpen: TMenuItem;
    StatusBar: TStatusBar;
    WebBrowser: TWebBrowser;
    TrayIcon: TTrayIcon;
    GridPanel: TGridPanel;
    Image2: TImage;
    DoDirs: TCheckBox;
    Image3: TImage;
    AskDirs: TCheckBox;
    Image4: TImage;
    CopyButton: TBitBtn;
    Image5: TImage;
    CopyPaths: TCheckBox;
    CopySizes: TCheckBox;
    Image7: TImage;
    ShellExecute: TBitBtn;
    Image1: TImage;
    FileGrid: TStringGrid;
    SettingsPanel: TRelativePanel;
    LangBox: TComboBox;
    ThemeBox: TComboBox;
    LangLabel: TLabel;
    ThemeLabel: TLabel;
    WarnText: TStaticText;
    HideSettingsButton: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure mSaveClick(Sender: TObject);
    procedure mNewClick(Sender: TObject);
    procedure mQuitClick(Sender: TObject);
    procedure FileGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mCopyAllClick(Sender: TObject);
    procedure mCopyNamesClick(Sender: TObject);
    procedure CopyButtonClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mmOpenClick(Sender: TObject);
    procedure mDeleteClick(Sender: TObject);
    procedure ShellExecuteClick(Sender: TObject);
    procedure msgTimerTimer(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ListBoxReload(Sender: TObject);
    procedure GridRenumerate(Sender: TObject);
    procedure ThemeBoxSelect(Sender: TObject);
    procedure SetStyle(style: String);
    procedure LoadIni(Sender: TObject);
    procedure SaveIni(Sender: TObject);
    procedure LoadLangIni(lang: string);
    procedure SaveLangIni(lang: string);
    procedure LangBoxChange(Sender: TObject);
    procedure LoadHelp(lang: String);
    procedure StatusBarDblClick(Sender: TObject);
    procedure SettingsPanelResize(Sender: TObject);

  private
    procedure WMDROPFILES(var Message: TWMDROPFILES); message WM_DROPFILES;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SymLinkForm: TSymLinkForm;
  // Lang variables
  MaxDragFiles, curNumFiles: Integer;
  sFiles, sList, sHelp, sSize, sFile, FolderStr, NotAFolder, srtUnavail, sPath, sMsgInclideFiles, sMsgDlgCaption, eFromEncode, eToEncode, sMaxDragFiles, sHelpFile, slastFileName: String;
  sNoData, sDoDirs, sDoDirsHint, sAskDirs, sAskDirsHint, sCopyButton, sCopyButtonHint, sCopySizes, sCopySizesHint, sCopyPaths, sCopyPathsHint, sShellExecute, sShellExecuteHint, mFile, mFileHint: string;
  // MenuNames lang variables
  sNeedFiles, smFile, smFileHint, smOpen, smOpenHint, smSave, smSaveHint, smCopy, smCopyHint, smNew, smNewHint, smQuit, smQuitHint, smFileList, smFileListHint, smDelete, smDeleteHint, smClear, smClearHint, smCopyAll, smCopyAllHint, smCopyNames,
    smCopyNamesHint, smpNew, smpNewHint, smpOpen, smpOpenHint, smpSave, smpSaveHint, smpCopyAll, smpCopyAllHint, smpCopyNames, smpCopyNamesHint, smpDelete, smpDeleteHint, smpQuit, smpQuitHint: string;
  // name of load lang
  currentLang, currentTheme: String;
  // Settings Windows
  sLangLabel, sLangBoxHint, sThemeLabel, sThemeBoxHint, sWarnText, sHideSettingsButton, sHideSettingsButtonHint, sStatusBarHint: String;

implementation

uses
  ShellApi, IniFiles, Vcl.Themes;

Var
  All_sizes: INT64;
  All_Files: Integer;
  str1, str2, commString: String;
  logFilePath, logFileName, logFileExt: string;
  longTimeStamp, hourTimeStamp: string;
  logFile: TextFile;
  batFile: TextFile;
  batFilePath, batFileName, batFileExt: string;
  commandLine: String;
  mklinkPathSlash: String;
  errorExecuting: String;

{$R *.dfm}
procedure AddToList(FileName: string); forward;

procedure TSymLinkForm.LoadIni(Sender: TObject);
var
  Ini: TIniFile;
begin
  // timeStamps
  longTimeStamp := FormatDateTime('yyyymmdd-hhnnsszzz', now);
  hourTimeStamp := FormatDateTime('yyyymmdd-hh', now);
  Ini := TIniFile.Create(Extractfilepath(paramstr(0)) + 'SymLinkCreator.ini');
  SymLinkForm.Width := Ini.ReadInteger('WinSize', 'Width', 1000);
  SymLinkForm.Height := Ini.ReadInteger('WinSize', 'Height', 680);
  SymLinkForm.Left := Ini.ReadInteger('WinPosition', 'X', 100);
  SymLinkForm.Top := Ini.ReadInteger('WinPosition', 'Y', 100);
  DoDirs.Checked := Ini.ReadBool('DoDirs', 'Checked', false);
  AskDirs.Checked := Ini.ReadBool('AskDirs', 'Checked', false);
  CopySizes.Checked := Ini.ReadBool('CopySizes', 'Checked', false);
  CopyPaths.Checked := Ini.ReadBool('CopyPaths', 'Checked', false);
  SymLinkForm.FileGrid.ColWidths[1] := Ini.ReadInteger('FileGridColWidth', 'Width1', 430);
  SymLinkForm.FileGrid.ColWidths[2] := Ini.ReadInteger('FileGridColWidth', 'Width2', 390);
  SymLinkForm.FileGrid.ColWidths[3] := Ini.ReadInteger('FileGridColWidth', 'Width3', 85);
  MaxDragFiles := Ini.ReadInteger('MaxDragFiles', 'Files', 25);
  sHelpFile := Ini.ReadString('Help', 'File', 'SimLinkCreatorHelpEn.html');
  TabNotebook.ActivePage := Ini.ReadString('Pages', 'ActivePage', 'Help');
  logFilePath := Ini.ReadString('Logs', 'fPath', Extractfilepath(paramstr(0)) + 'Logs');
  logFileName := Ini.ReadString('Logs', 'fName', 'SimLinkCreator');
  logFileExt := Ini.ReadString('Logs', 'fExt', '.log');
  batFilePath := Ini.ReadString('Bat', 'fPath', Extractfilepath(paramstr(0)) + 'Bat');
  batFileName := Ini.ReadString('Bat', 'fName', 'SimLinkCreator');
  batFileExt := Ini.ReadString('Bat', 'fExt', '.bat');
  slastFileName := Ini.ReadString('LastFile', 'File', Extractfilepath(paramstr(0)) + 'SymLinkCreator_last.txt');
  currentLang := Ini.ReadString('Lang', 'Lang', 'en');
  currentTheme := Ini.ReadString('Theme', 'Theme', 'Carbon');
  Ini.Free;
end;

procedure TSymLinkForm.SaveIni(Sender: TObject);
var
  Ini: TIniFile;
begin
  // Save INI-file
  Ini := TIniFile.Create(Extractfilepath(paramstr(0)) + 'SymLinkCreator.ini');
  Ini.WriteInteger('WinSize', 'Width', SymLinkForm.Width);
  Ini.WriteInteger('WinSize', 'Height', SymLinkForm.Height);
  Ini.WriteString('WinPosition', 'Comment', 'WinPos Stability only with "Windows 10" theme!');
  Ini.WriteInteger('WinPosition', 'X', SymLinkForm.Left);
  Ini.WriteInteger('WinPosition', 'Y', SymLinkForm.Top);
  Ini.WriteInteger('FileGridColWidth', 'Width1', FileGrid.ColWidths[1]);
  Ini.WriteInteger('FileGridColWidth', 'Width2', FileGrid.ColWidths[2]);
  Ini.WriteInteger('FileGridColWidth', 'Width3', FileGrid.ColWidths[3]);
  Ini.WriteInteger('MaxDragFiles', 'Files', MaxDragFiles);
  Ini.WriteString('Help', 'File', sHelpFile);
  Ini.WriteString('Pages', 'ActivePAge', TabNotebook.ActivePage);
  Ini.WriteString('Logs', 'fPath', logFilePath);
  Ini.WriteString('Logs', 'fName', logFileName);
  Ini.WriteString('Logs', 'fExt', logFileExt);
  Ini.WriteString('Bat', 'fPath', batFilePath);
  Ini.WriteString('Bat', 'fName', batFileName);
  Ini.WriteString('Bat', 'fExt', batFileExt);
  Ini.WriteString('LastFile', 'File', slastFileName);
  Ini.WriteBool('DoDirs', 'Checked', DoDirs.Checked);
  Ini.WriteBool('AskDirs', 'Checked', AskDirs.Checked);
  Ini.WriteBool('CopySizes', 'Checked', CopySizes.Checked);
  Ini.WriteBool('CopyPaths', 'Checked', CopyPaths.Checked);
  Ini.WriteString('Lang', 'Lang', currentLang);
  Ini.WriteString('Theme', 'Theme', currentTheme);
  Ini.Free;
end;

procedure TSymLinkForm.LoadLangIni(lang: String);
var
  langIni: TIniFile;
begin
  // open lang files
  CreateDir(Extractfilepath(paramstr(0)) + 'Lang');
  langIni := TIniFile.Create(Extractfilepath(paramstr(0)) + 'Lang\' + currentLang + '.ini');
  sNeedFiles := langIni.ReadString('Lang', 'NeedFiles', 'Drag files/folders then start "Create links"');
  FolderStr := langIni.ReadString('Lang', 'FolderStr', 'Folder');
  NotAFolder := langIni.ReadString('Lang', 'folderFirts', 'First string must be FOLDER!!!');
  srtUnavail := langIni.ReadString('Lang', 'Unavailable', 'Unavailable');
  sPath := langIni.ReadString('Lang', 'Path', 'Path');
  sFile := langIni.ReadString('Lang', 'File', 'File');
  sSize := langIni.ReadString('Lang', 'Sise', 'Size');
  sFiles := langIni.ReadString('Lang', 'TabFiles', 'Files');
  sList := langIni.ReadString('Lang', 'TabList', 'Flat list');
  sHelp := langIni.ReadString('Lang', 'TabHelp', 'Help');
  sMsgInclideFiles := langIni.ReadString('Lang', 'sMsgInclideFiles', 'Include files in them instead of folders?');
  sMsgDlgCaption := langIni.ReadString('Lang', 'sMsgDlgCaption', 'The logic of enabling folders');
  srtUnavail := langIni.ReadString('Lang', 'srtUnavail', 'Unavailable');
  eFromEncode := langIni.ReadString('Lang', 'eFromEncode', '1251');
  eToEncode := langIni.ReadString('Lang', 'eToEncode', '866');
  sMaxDragFiles := langIni.ReadString('Lang', 'sMaxDragFiles', ' ' + IntToStr(MaxDragFiles) + ' files max - ini-file settings');
  sNoData := langIni.ReadString('Lang', 'NoData', 'No data');
  sDoDirs := langIni.ReadString('Lang', 'DoDirs', 'files instead of folders or >');
  sDoDirsHint := langIni.ReadString('Lang', 'DoDirsHint', 'Include files in folders instead of folders in the list');
  sDoDirs := langIni.ReadString('Lang', 'DoDirs', 'files instead of folders or >');
  sAskDirs := langIni.ReadString('Lang', 'AskDirsText', 'ask all time');
  sAskDirsHint := langIni.ReadString('Lang', 'AskDirsHint', 'Ask everyone about the folders in the list');
  sCopyButton := langIni.ReadString('Lang', 'CopyButton', 'Copy file list');
  sCopyButtonHint := langIni.ReadString('Lang', 'CopyButtonHint', 'Copy file list to buffer');
  sCopySizes := langIni.ReadString('Lang', 'CopySizes', '+ size');
  sCopySizesHint := langIni.ReadString('Lang', 'CopySizesHint', 'Add the file size to the file names as well');
  sCopyPaths := langIni.ReadString('Lang', 'CopyPaths', '+ path');
  sCopyPathsHint := langIni.ReadString('Lang', 'CopyPathsHint', 'Add the full paths to the files');
  sShellExecute := langIni.ReadString('Lang', 'ShellExecute', 'Create Links');
  sShellExecuteHint := langIni.ReadString('Lang', 'ShellExecuteHint', 'Run script and create symbolic links');
  // load all menu lang strings
  smFile := langIni.ReadString('Lang-Menu', 'smFile', 'File');
  smFileHint := langIni.ReadString('Lang-Menu', 'smFileHint', 'File operations');
  smOpen := langIni.ReadString('Lang-Menu', 'smOpen', 'Open');
  smOpenHint := langIni.ReadString('Lang-Menu', 'smOpenHint', 'Open saved file list');
  smSave := langIni.ReadString('Lang-Menu', 'smSave', 'Save');
  smSaveHint := langIni.ReadString('Lang-Menu', 'smSaveHint', 'Save file list');
  smNew := langIni.ReadString('Lang-Menu', 'smNew', 'New');
  smNewHint := langIni.ReadString('Lang-Menu', 'smNewHint', 'New file list');
  smQuit := langIni.ReadString('Lang-Menu', 'smQuit', 'Quit');
  smQuitHint := langIni.ReadString('Lang-Menu', 'smQuitHint', 'Finish and exit');
  smFileList := langIni.ReadString('Lang-Menu', 'smFileList', 'File list');
  smFileListHint := langIni.ReadString('Lang-Menu', 'smFileListHint', 'Operations with filelist');
  smDelete := langIni.ReadString('Lang-Menu', 'smDelete', 'Delete');
  smDeleteHint := langIni.ReadString('Lang-Menu', 'smDeleteHint', 'Delete current string');
  smClear := langIni.ReadString('Lang-Menu', 'smClear', 'Clear');
  smClearHint := langIni.ReadString('Lang-Menu', 'smClearHint', 'Clear all');
  smCopyAll := langIni.ReadString('Lang-Menu', 'smCopyAll', 'Copy all');
  smCopyAllHint := langIni.ReadString('Lang-Menu', 'smCopyAllHint', 'Copy all to buffer');
  smCopyNames := langIni.ReadString('Lang-Menu', 'smCopyNames', 'Copy only names');
  smCopyNamesHint := langIni.ReadString('Lang-Menu', 'smCopyNamesHint', 'Copy only names to buffer');
  smpNew := langIni.ReadString('Lang-Menu', 'smpNew', 'New');
  smpOpen := langIni.ReadString('Lang-Menu', 'smpOpen', 'Open');
  smpSave := langIni.ReadString('Lang-Menu', 'smpSave', 'Save');
  smpCopyAll := langIni.ReadString('Lang-Menu', 'smpCopyAll', 'Copy all');
  smpCopyNames := langIni.ReadString('Lang-Menu', 'smpCopyNames', 'Copy names');
  smpDelete := langIni.ReadString('Lang-Menu', 'smpDelete', 'Delete');
  smpQuit := langIni.ReadString('Lang-Menu', 'smpQuit', 'Quit');
  smpNew := langIni.ReadString('Lang-Menu', 'mNew', 'New');
  smpNewHint := langIni.ReadString('Lang-Menu', 'mNewHint', 'New file-list');
  smpOpen := langIni.ReadString('Lang-Menu', 'mOpen', 'Open');
  smpOpenHint := langIni.ReadString('Lang-Menu', 'mOpenHint', 'Open file-list');
  smpSave := langIni.ReadString('Lang-Menu', 'mSave', 'Save');
  smpSaveHint := langIni.ReadString('Lang-Menu', 'mSaveHint', 'Save file-list');
  smpCopyAll := langIni.ReadString('Lang-Menu', 'smpCopyAll', 'Copy all - 2');
  smpCopyAllHint := langIni.ReadString('Lang-Menu', 'smpCopyHint', 'Copy all to buffer');
  smpCopyNames := langIni.ReadString('Lang-Menu', 'mCopyNames', 'Copy names');
  smpCopyNamesHint := langIni.ReadString('Lang-Menu', 'mCopyNamesHint', 'Copy only file names');
  smpDelete := langIni.ReadString('Lang-Menu', 'mDelete', 'Delete');
  smpDeleteHint := langIni.ReadString('Lang-Menu', 'mDeleteHint', 'Delete current string');
  smpQuit := langIni.ReadString('Lang-Menu', 'mQuit', 'Quit');
  smpQuitHint := langIni.ReadString('Lang-Menu', 'mQuitHint', 'Exit program');
  sLangLabel := langIni.ReadString('Settings', 'LangLabel', 'Language');
  sLangBoxHint := langIni.ReadString('Settings', 'LangBoxHint', 'Select a language from the list');
  sThemeLabel := langIni.ReadString('Settings', 'ThemeLabel', 'Theme');
  sThemeBoxHint := langIni.ReadString('Settings', 'ThemeBoxHint', 'elect Theme from the list');
  sWarnText := langIni.ReadString('Settings', 'WarnText', 'Save your data! These settings will reload the interface!');
  sHideSettingsButton := langIni.ReadString('Settings', 'HideSettingsButton', 'Exit without saving');
  sHideSettingsButtonHint := langIni.ReadString('Settings', 'HideSettingsButtonHint', 'Exit without change settings');
  sStatusBarHint := langIni.ReadString('Settings', 'StatusBarHint', 'Double click opens settings');
  langIni.Free;
  // init main menu lang strings
  MainMenu.Items.Items[0].Caption := smFile;
  MainMenu.Items.Items[0].Hint := smFileHint;
  MainMenu.Items.Items[0].Items[0].Caption := smOpen;
  MainMenu.Items.Items[0].Items[0].Hint := smOpenHint;
  MainMenu.Items.Items[0].Items[1].Caption := smSave;
  MainMenu.Items.Items[0].Items[1].Hint := smSaveHint;
  MainMenu.Items.Items[0].Items[2].Caption := smNew;
  MainMenu.Items.Items[0].Items[2].Hint := smNewHint;
  MainMenu.Items.Items[0].Items[3].Caption := smQuit;
  MainMenu.Items.Items[0].Items[3].Hint := smQuitHint;
  // second main menu
  MainMenu.Items.Items[1].Caption := smFileList;
  MainMenu.Items.Items[1].Hint := smFileListHint;
  MainMenu.Items.Items[1].Items[0].Caption := smDelete;
  MainMenu.Items.Items[1].Items[0].Hint := smDeleteHint;
  MainMenu.Items.Items[1].Items[1].Caption := smClear;
  MainMenu.Items.Items[1].Items[1].Hint := smClearHint;
  MainMenu.Items.Items[1].Items[2].Caption := smCopyAll;
  MainMenu.Items.Items[1].Items[2].Hint := smCopyAllHint;
  MainMenu.Items.Items[1].Items[3].Caption := smCopyNames;
  MainMenu.Items.Items[1].Items[3].Hint := smCopyNamesHint;
  // init pop-up menu lang strings
  pMenu.Items[0].Caption := smpNew;
  pMenu.Items[0].Hint := smpNewHint;
  pMenu.Items[1].Caption := smpOpen;
  pMenu.Items[1].Hint := smpOpenHint;
  pMenu.Items[2].Caption := smpSave;
  pMenu.Items[2].Hint := smpSaveHint;
  pMenu.Items[3].Caption := smpCopyAll;
  pMenu.Items[3].Hint := smpCopyAllHint;
  pMenu.Items[4].Caption := smpCopyNames;
  pMenu.Items[4].Hint := smpCopyNamesHint;
  pMenu.Items[5].Caption := smpDelete;
  pMenu.Items[5].Hint := smpDeleteHint;
  pMenu.Items[6].Caption := smpQuit;
  pMenu.Items[6].Hint := smpQuitHint;
  // init tools lang strings
  ShellExecute.Caption := sShellExecute;
  ShellExecute.Hint := sShellExecuteHint;
  CopyPaths.Caption := sCopyPaths;
  CopyPaths.Hint := sCopyPathsHint;
  CopySizes.Caption := sCopySizes;
  CopySizes.Hint := sCopySizesHint;
  CopyButton.Caption := sCopyButton;
  CopyButton.Hint := sCopyButtonHint;
  AskDirs.Caption := sAskDirs;
  AskDirs.Hint := sAskDirsHint;
  DoDirs.Hint := sDoDirsHint;
  DoDirs.Caption := sDoDirs;
  StatusBar.SimpleText := sNoData;
  FileGrid.Cells[0, 0] := '*';
  FileGrid.Cells[1, 0] := sPath;
  FileGrid.Cells[2, 0] := sFile;
  FileGrid.Cells[3, 0] := sSize;
  // Names of tabs in notebook
  TabNotebook.Pages[0] := sFiles;
  TabNotebook.Pages[1] := sList;
  TabNotebook.Pages[2] := sHelp;
  // Settings windows initial
  LangLabel.Caption := sLangLabel;
  LangBox.Hint := sLangBoxHint;
  ThemeLabel.Caption := sThemeLabel;
  ThemeBox.Hint := sThemeBoxHint;
  WarnText.Caption := sWarnText;
  HideSettingsButton.Caption := sHideSettingsButton;
  HideSettingsButton.Hint := sHideSettingsButtonHint;
  StatusBar.Hint := sStatusBarHint;
end;

procedure TSymLinkForm.SaveLangIni(lang: String);
var
  langIni: TIniFile;
begin
  // Save main lang strings
  if currentLang = '12345' then
  begin
    langIni := TIniFile.Create(Extractfilepath(paramstr(0)) + 'Lang\' + currentLang + '.ini');
    langIni.WriteString('Lang', 'FolderStr', FolderStr);
    langIni.WriteString('Lang', 'NeedFiles', sNeedFiles);
    langIni.WriteString('Lang', 'folderFirts', NotAFolder);
    langIni.WriteString('Lang', 'Unavailable', srtUnavail);
    langIni.WriteString('Lang', 'Path', sPath);
    langIni.WriteString('Lang', 'File', sFile);
    langIni.WriteString('Lang', 'Sise', sSize);
    langIni.WriteString('Lang', 'TabFiles', sFiles);
    langIni.WriteString('Lang', 'TabList', sList);
    langIni.WriteString('Lang', 'TabHelp', sHelp);
    langIni.WriteString('Lang', 'sMsgInclideFiles', sMsgInclideFiles);
    langIni.WriteString('Lang', 'sMsgDlgCaption', sMsgDlgCaption);
    langIni.WriteString('Lang', 'srtUnavail', srtUnavail);
    langIni.WriteString('Lang', 'eFromEncode', eFromEncode);
    langIni.WriteString('Lang', 'eToEncode', eToEncode);
    langIni.WriteString('Lang', 'sMaxDragFiles', sMaxDragFiles);
    langIni.WriteString('Lang', 'NoData', sNoData);
    langIni.WriteString('Lang', 'DoDirs', sDoDirs);
    langIni.WriteString('Lang', 'DoDirsHint', sDoDirsHint);
    langIni.WriteString('Lang', 'CopyButton', sCopyButton);
    langIni.WriteString('Lang', 'CopyButtonHint', sCopyButtonHint);
    langIni.WriteString('Lang', 'CopySizes', sCopySizes);
    langIni.WriteString('Lang', 'CopySizesHint', sCopySizesHint);
    langIni.WriteString('Lang', 'CopyPaths', sCopyPaths);
    langIni.WriteString('Lang', 'CopyPathsHint', sCopyPathsHint);
    langIni.WriteString('Lang', 'ShellExecute', sShellExecute);
    langIni.WriteString('Lang', 'ShellExecuteHint', sShellExecuteHint);
    langIni.WriteString('Lang', 'AskDirsText', sAskDirs);
    langIni.WriteString('Lang', 'AskDirsHint', sAskDirsHint);
    // Save menu lang strings
    langIni.WriteString('Lang-Menu', 'smFile', smFile);
    langIni.WriteString('Lang-Menu', 'smFileHint', smFileHint);
    langIni.WriteString('Lang-Menu', 'smOpen', smOpen);
    langIni.WriteString('Lang-Menu', 'smOpenHint', smOpenHint);
    langIni.WriteString('Lang-Menu', 'smSave', smSave);
    langIni.WriteString('Lang-Menu', 'smSaveHint', smSaveHint);
    langIni.WriteString('Lang-Menu', 'smNew', smNew);
    langIni.WriteString('Lang-Menu', 'smNewHint', smNewHint);
    langIni.WriteString('Lang-Menu', 'smQuit', smQuit);
    langIni.WriteString('Lang-Menu', 'smQuitHint', smQuitHint);
    langIni.WriteString('Lang-Menu', 'smFileList', smFileList);
    langIni.WriteString('Lang-Menu', 'smFileListHint', smFileListHint);
    langIni.WriteString('Lang-Menu', 'smDelete', smDelete);
    langIni.WriteString('Lang-Menu', 'smDeleteHint', smDeleteHint);
    langIni.WriteString('Lang-Menu', 'smClear', smClear);
    langIni.WriteString('Lang-Menu', 'smClearHint', smClearHint);
    langIni.WriteString('Lang-Menu', 'smCopyAll', smCopyAll);
    langIni.WriteString('Lang-Menu', 'smCopyAllHint', smCopyAllHint);
    langIni.WriteString('Lang-Menu', 'smCopyNames', smCopyNames);
    langIni.WriteString('Lang-Menu', 'smCopyNamesHint', smCopyNamesHint);
    langIni.WriteString('Lang-Menu', 'mNew', smNew);
    langIni.WriteString('Lang-Menu', 'mNewHint', smNewHint);
    langIni.WriteString('Lang-Menu', 'mOpen', smOpen);
    langIni.WriteString('Lang-Menu', 'mOpenHint', smOpenHint);
    langIni.WriteString('Lang-Menu', 'mSave', smSave);
    langIni.WriteString('Lang-Menu', 'mSaveHint', smSaveHint);
    langIni.WriteString('Lang-Menu', 'mCopyNames', smCopyNames);
    langIni.WriteString('Lang-Menu', 'mCopyNamesHint', smCopyNamesHint);
    langIni.WriteString('Lang-Menu', 'mDelete', smDelete);
    langIni.WriteString('Lang-Menu', 'mDeleteHint', smDeleteHint);
    langIni.WriteString('Lang-Menu', 'mQuit', smQuit);
    langIni.WriteString('Lang-Menu', 'mQuitHint', smQuitHint);
    langIni.WriteString('Lang-Menu', 'smpNew', smpNew);
    langIni.WriteString('Lang-Menu', 'smpNewHint', smpNewHint);
    langIni.WriteString('Lang-Menu', 'smpOpen', smpOpen);
    langIni.WriteString('Lang-Menu', 'smpOpenHint', smpOpenHint);
    langIni.WriteString('Lang-Menu', 'smpSave', smpSave);
    langIni.WriteString('Lang-Menu', 'smpSaveHint', smpSaveHint);
    langIni.WriteString('Lang-Menu', 'smpCopyAll', smpCopyAll);
    langIni.WriteString('Lang-Menu', 'smpCopyAllHint', smpCopyAllHint);
    langIni.WriteString('Lang-Menu', 'smpCopyNames', smpCopyNames);
    langIni.WriteString('Lang-Menu', 'smpCopyNamesHint', smpCopyNamesHint);
    langIni.WriteString('Lang-Menu', 'smpDelete', smpDelete);
    langIni.WriteString('Lang-Menu', 'smpDeleteHint', smpDeleteHint);
    langIni.WriteString('Lang-Menu', 'smpQuit', smpQuit);
    langIni.WriteString('Lang-Menu', 'smpQuitHint', smpQuitHint);
    langIni.WriteString('Settings', 'LangLabel', sLangLabel);
    langIni.WriteString('Settings', 'LangBoxHint', sLangBoxHint);
    langIni.WriteString('Settings', 'ThemeLabel', sThemeLabel);
    langIni.WriteString('Settings', 'ThemeBoxHint', sThemeBoxHint);
    langIni.WriteString('Settings', 'WarnText', sWarnText);
    langIni.WriteString('Settings', 'HideSettingsButton', sHideSettingsButton);
    langIni.WriteString('Settings', 'HideSettingsButtonHint', sHideSettingsButtonHint);
    langIni.WriteString('Settings', 'StatusBarHint', sStatusBarHint);
    langIni.Free;
  end;
end;

procedure TSymLinkForm.LoadHelp(lang: String);
begin
  // try load Help page without exceptions
  try
    if FileExists(sHelpFile) then
      WebBrowser.Navigate(Extractfilepath(paramstr(0)) + sHelpFile)
    else
      TabNotebook.Pages.Delete(2);
  finally
  end;
end;

/// ///////////////////////////////////////////////////////////// START FormCreate
procedure TSymLinkForm.FormCreate(Sender: TObject);
var
  i, a: Integer;
  Ini, langIni: TIniFile;
  List: TStringList;
  f: TextFile;
  s: String;
  sm: TStyleManager;
  si: TStyleInfo;
  FileArray, LangArray: TStringDynArray;
  m: TMenuItem;
begin
  SymLinkForm.Canvas.Font.Charset := GB2312_CHARSET;
  LoadIni(Self); // load ini-files
  LoadLangIni(currentLang);
  LoadHelp(currentLang);
  // try load last file-list without exceptions
  try
    if FileExists(slastFileName) then
    begin
      ListBox.Clear;
      mNewClick(Self);
      AssignFile(f, slastFileName);
      Reset(f);
      a := 0;
      while (not EOF(f)) do
      begin
        Readln(f, s);
        if s = '' then
          break;
        AddToList(s);
        Inc(a);
      end;
      CloseFile(f)
    end;
  finally
  end;
  sm := TStyleManager.Create; // build list of available Visual Styles
  begin
    CreateDir(Extractfilepath(paramstr(0)) + 'Styles');
    FileArray := TDirectory.GetFiles(Extractfilepath(paramstr(0)) + 'Styles\', '*.vsf'); // Find vsf-files
    // check & validate founded styles
    for i := 0 to Length(FileArray) - 1 do
    begin
      if TStyleManager.IsValidStyle(FileArray[i]) then
      begin
        if TStyleManager.style[si.Name] = nil then
        begin
          TStyleManager.LoadFromFile(FileArray[i]);
        end;
      end;
    end;
    for i := 0 to Length(sm.StyleNames) - 1 do
      ThemeBox.Items.Add(sm.StyleNames[i]); // Add list available styles
    ThemeBox.Sorted := true;
    Try // Try losd Style, if error - use built in Windows style
      TStyleManager.TrySetStyle(currentTheme);
    Finally
    End;
    Application.ProcessMessages; // Restore Drag&Drop accept to new Form ID if its change
    DragAcceptFiles(SymLinkForm.Handle, true); // accept drag files to program
  end;

  try
    CreateDir(Extractfilepath(paramstr(0)) + 'Lang');
    LangArray := TDirectory.GetFiles(Extractfilepath(paramstr(0)) + 'Lang\', '*.ini'); // Find lang files}
    for i := 0 to Length(LangArray) - 1 do
    begin
      LangBox.Items.Add(TPath.GetFileNameWithoutExtension(LangArray[i]));
    end;
  finally
  end;

end;

/// ///////////////////////////////////////////////////////////// END FormCreate
procedure TSymLinkForm.LangBoxChange(Sender: TObject);
begin
  currentLang := LangBox.Items[LangBox.ItemIndex];
  LoadLangIni(currentLang);
end;

procedure TSymLinkForm.SetStyle(style: String);
var
  si: TStyleInfo;
  s: String;
begin
  if TStyleManager.IsValidStyle(s, si) then
  begin
    // load stile
    if TStyleManager.style[si.Name] = nil then
    begin
      TStyleManager.LoadFromFile(s);
    end;
  end
  else;
end;

procedure TSymLinkForm.SettingsPanelResize(Sender: TObject);
begin
  SettingsPanel.Left := ((SymLinkForm.Width div 2) - (SettingsPanel.Width div 2)) - 11;
  SettingsPanel.Top := ((SymLinkForm.Height div 2) - ((SettingsPanel.Height div 2)) - 75);
end;

// change Style
procedure TSymLinkForm.ThemeBoxSelect(Sender: TObject);
begin
  TStyleManager.TrySetStyle(ThemeBox.Text, false);
  currentTheme := ThemeBox.Text;
  Application.ProcessMessages; // process the message queue;
  DragAcceptFiles(SymLinkForm.Handle, true); // restore Dreag&Drop accept
end;

procedure TSymLinkForm.FormDestroy(Sender: TObject);
var
  a: Integer;
  f: TextFile;
  str: string;
begin
  SaveIni(Self);
  SaveLangIni(currentLang);
  // Load last fileList
  begin
    ListBoxReload(Self);
    AssignFile(f, slastFileName);
    Rewrite(f);
    With ListBox do
    begin
      for a := 0 to Count - 1 do
      begin
        str := ListBox.Items[a];
        Writeln(f, str);
      end;
    end;
    CloseFile(f);
  end;
end;

procedure InitLang(lang: String);
var
  langIni: TIniFile;
begin
end;

// Encoder
function WinToDos(const s: string): string;
var
  b: TBytes;
begin
  b := BytesOf(s);
  b := TEncoding.Convert(TEncoding.GetEncoding(StrToInt(eFromEncode)), TEncoding.GetEncoding(StrToInt(eToEncode)), b);
  Result := StringOf(b);
end;

function utf8To1251(const s: string): string;
var
  b: TBytes;
begin
  b := BytesOf(s);
  b := TEncoding.Convert(TEncoding.GetEncoding(StrToInt('utf8')), TEncoding.GetEncoding(StrToInt('1251')), b);
  Result := StringOf(b);
end;

procedure findlastmath(str1, str2: string; var findtopos: Integer);
var
  a: Integer;
begin
  if Length(str1) < findtopos then
    findtopos := Length(str1);
  if Length(str2) < findtopos then
    findtopos := Length(str2);
  For a := 1 to findtopos do
    if str1[a] <> str2[a] then
      break;
  findtopos := a - 1;
end;

Procedure AddToList(FileName: string);
  Procedure AddToNewRow(path, nam, size: string);
  var
    Rowtoadd: Integer;
  begin
    With SymLinkForm do
    begin
      If (FileGrid.RowCount = 1) then
        FileGrid.RowCount := FileGrid.RowCount + 1;
      If (FileGrid.RowCount = 2) and (FileGrid.Cells[0, 1] = '') then
        Rowtoadd := 1
      else
      begin
        Rowtoadd := FileGrid.RowCount;
        if Rowtoadd > MaxDragFiles then
        begin
          StatusBar.SimpleText := sMaxDragFiles;
          ShellExecute.Enabled := false;
          exit;
        end;
        FileGrid.RowCount := FileGrid.RowCount + 1;
      end;
      FileGrid.Cells[0, Rowtoadd] := IntToStr(Rowtoadd);
      FileGrid.Cells[1, Rowtoadd] := path;
      FileGrid.Cells[2, Rowtoadd] := nam;
      FileGrid.Cells[3, Rowtoadd] := size;
      ListBox.Items.Add(path + nam);
    end;
  end;

  Procedure ScanDir(path: string);
  var
    srs: TSearchRec;
    filefound: boolean;
  begin
    Application.ProcessMessages;
    filefound := false;
    if FindFirst(path + '\*.*', faAnyFile, srs) = 0 then
    begin
      While findnext(srs) = 0 do
      begin
        if srs.Name = '..' then
          continue;
        filefound := true;
        if (srs.Attr and faDirectory) <> 0 then // scan subfsssss....
          ScanDir(path + '\' + srs.Name)
        ELSE // work as file
        begin
          AddToNewRow(path + '\', srs.Name, IntToStr(srs.size));
          All_sizes := All_sizes + srs.size;
          Inc(All_Files);
        end;
      end; // cycle FindNext
    end;
    findclose(srs);
    if not filefound then
      AddToNewRow(Extractfilepath(path), Extractfilename(path), FolderStr);
  end;

var
  sr: TSearchRec;
begin // of Addtolist(Filename:string);
  With SymLinkForm do
  begin
    if FindFirst(FileName, faAnyFile, sr) = 0 then
    begin
      if (sr.Attr and faDirectory) <> 0 then
      begin // adding folder to list
        If DoDirs.Checked then
          ScanDir(FileName)
        else // adding as folder
          AddToNewRow(Extractfilepath(FileName), Extractfilename(FileName), FolderStr);
      end
      else
      Begin // addinf file to list
        AddToNewRow(Extractfilepath(FileName), Extractfilename(FileName), IntToStr(sr.size));
        All_sizes := All_sizes + sr.size;
        Inc(All_Files);
      end;
    end
    else // any error... may be deleted file while dropping ...
      AddToNewRow(Extractfilepath(FileName), Extractfilename(FileName), srtUnavail);
    findclose(sr);
  end; { With ... }
end;

procedure TSymLinkForm.WMDROPFILES(var Message: TWMDROPFILES);
var
  NumFiles: longint;
  i: longint;
  buffer: array [0 .. 1024] of char;
begin
  If AskDirs.Checked then // How many files are dropped
  begin
    DoDirs.Checked := Application.MessageBox(PChar(sMsgInclideFiles), PChar(sMsgDlgCaption), MB_YESNO) = IDYes;
  end;
  NumFiles := DragQueryFile(Message.Drop, cardinal(-1), nil, 0);
  for i := 0 to (NumFiles - 1) do // Accept dropped files
  begin
    DragQueryFile(Message.Drop, i, @buffer, sizeof(buffer));
    AddToList(buffer);
  end;
end;

procedure appendLogFile(LogString: string);
begin
  If NOT DirectoryExists(logFilePath) then
    try
      CreateDir(logFilePath)
    finally
    end;

  if FileExists(logFilePath + '\' + hourTimeStamp + '-' + logFileName + logFileExt) then
  begin
    try
      AssignFile(logFile, logFilePath + '\' + hourTimeStamp + '-' + logFileName + logFileExt);
      Append(logFile);
      Writeln(logFile, LogString);
      CloseFile(logFile);
    finally
    end;
  end
  else
  begin
    AssignFile(logFile, logFilePath + '\' + hourTimeStamp + '-' + logFileName + logFileExt);
    Rewrite(logFile);
    Writeln(logFile, LogString);
    CloseFile(logFile);
  end;
end;

procedure saveBatFile(LogString: string);
var
  sList: TStringList;
  s, line1: String;
  a: Integer;
begin
  longTimeStamp := FormatDateTime('yyyymmdd-hhnnsszzz', now);
  hourTimeStamp := FormatDateTime('yyyymmdd-hh', now);
  If NOT DirectoryExists(batFilePath) then
    try
      CreateDir(batFilePath)
    finally
    end;
  commandLine := '';
  line1 := '';
  sList := TStringList.Create;
  line1 := 'mklink "' + SymLinkForm.ListBox.Items[0];
  For a := 2 to SymLinkForm.ListBox.Items.Count do
  begin
    line1 := 'mklink "' + SymLinkForm.ListBox.Items[0] + '\' + SymLinkForm.FileGrid.Cells[2, a] + '"';
    if SymLinkForm.FileGrid.Cells[3, a] = FolderStr then
      s := '/D '
    else
      s := '';
    sList.Add(WinToDos(line1 + ' ' + s + '"' + SymLinkForm.FileGrid.Cells[1, a] + SymLinkForm.FileGrid.Cells[2, a] + '"'));
  end;
  sList.SaveToFile(batFilePath + '\' + longTimeStamp + '-' + batFileName + batFileExt);
  sList.SaveToFile('temp.bat');
  try
    WinExec('temp.bat', SW_HIDE);
  finally
  end;
  // commandLine := batFilePath + '\' + longTimeStamp + '-' + batFileName + batFileExt;
  appendLogFile(hourTimeStamp + ': ' + batFilePath + '\' + longTimeStamp + '-' + batFileName + batFileExt);
  sList.Free;
  sleep(500);
  DeleteFile('temp.bat');
end;

procedure TSymLinkForm.ShellExecuteClick(Sender: TObject);
begin
  saveBatFile('');
  WinExec('test.cmd', SW_HIDE);
end;

procedure TSymLinkForm.StatusBarDblClick(Sender: TObject);
begin
  SettingsPanel.Visible := NOT SettingsPanel.Visible;
end;

procedure TSymLinkForm.mSaveClick(Sender: TObject);
Var
  f: TextFile;
  str: string;
  a: Integer;
begin
  If SaveFileDialog.Execute then
  begin
    ListBoxReload(Self);
    AssignFile(f, SaveFileDialog.FileName);
    Rewrite(f);
    With ListBox do
    begin
      for a := 0 to Count - 1 do
      begin
        str := ListBox.Items[a];
        Writeln(f, str);
      end;
    end;
    CloseFile(f);
  end;
end;

procedure TSymLinkForm.ListBoxReload(Sender: TObject);
var
  a: Integer;
  str: string;
begin
  ListBox.Items.Clear;
  for a := 1 to FileGrid.RowCount - 1 do
  begin
    Application.ProcessMessages;
    if FileGrid.Cells[1, a] = '' then
      break;
    str := (FileGrid.Cells[1, a] + FileGrid.Cells[2, a]);
    ListBox.Items.Add(str);
  end;
  for a := 1 to FileGrid.RowCount - 201 do
    FileGrid.Cells[0, a] := IntToStr(a);
end;

procedure TSymLinkForm.msgTimerTimer(Sender: TObject);
var
  strFolder: string;
begin
  strFolder := FileGrid.Cells[3, 1];
  if strFolder.Length = 0 then
  begin
    ShellExecute.Enabled := false;
    StatusBar.Panels.Items[0].Text := NotAFolder;
    exit;
  end;
  if strFolder <> FolderStr then
  begin
    ShellExecute.Enabled := false;
    StatusBar.Panels.Items[0].Text := NotAFolder;
    mNewClick(Self);
    exit;
  end;
  if (strFolder = FolderStr) and (ListBox.Items.Count = 1) then
  begin
    // StatusBar.Panels.Items[0].Text := FileGrid.Cells[3, 2];
    StatusBar.Panels.Items[0].Text := sNeedFiles;
    ShellExecute.Enabled := false;
    exit;
  end;
  if (strFolder = FolderStr) and (ListBox.Items.Count > 1) then
  begin
    StatusBar.Panels.Items[0].Text := sNeedFiles;
    ShellExecute.Enabled := true;
  end;
  if ListBox.Items.Count >= MaxDragFiles then
  begin
    StatusBar.Panels.Items[0].Text := sMaxDragFiles + ' - "' + sShellExecute + '"';
  end;
end;

procedure TSymLinkForm.mmOpenClick(Sender: TObject);
var
  f: TextFile;
  s: String;
  a: Integer;
begin
  if OpenFileDialog.Execute then
  begin
    ListBox.Clear;
    mNewClick(Self);
    AssignFile(f, OpenFileDialog.FileName);
    Reset(f);
    a := 0;
    while (not EOF(f)) do
    begin
      Readln(f, s);
      if s = '' then
        break;
      AddToList(s);
      Inc(a);
    end;
    CloseFile(f)
  end;
end;

procedure TSymLinkForm.CopyButtonClick(Sender: TObject);
Var
  buff: PChar;
  st: string;
  a: Integer;
  strs: Tstrings;
begin
  strs := TStringList.Create;
  try
    for a := 1 to FileGrid.RowCount - 1 do
    begin
      st := '';
      If CopyPaths.Checked then
        st := FileGrid.Cells[1, a] + FileGrid.Cells[2, a]
      else
        st := FileGrid.Cells[2, a];
      If CopySizes.Checked then
        st := st + #09 + FileGrid.Cells[3, a] { + #13 + #10 }
      else
        st := st { + #13 + #10 };
      strs.Add(st);
    end;
    buff := PChar(strs.Text);
    Clipboard.SetTextBuf(buff);
  finally
    strs.Free;
  end;

end;

procedure TSymLinkForm.GridRenumerate(Sender: TObject);
var
  a: Integer;
begin
  With SymLinkForm do
  begin
    for a := 1 to MaxDragFiles do
    begin
      if FileGrid.Cells[1, a] = '' then
        exit;
      FileGrid.Cells[0, a] := IntToStr(a);
    end;
  end;
end;

procedure TSymLinkForm.FileGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var
  a: Integer;
begin
  if Key = 46 { del } then
  begin
    for a := FileGrid.Row to FileGrid.RowCount - 2 do
      FileGrid.Rows[a] := FileGrid.Rows[a + 1];
    if FileGrid.RowCount > 2 then
      FileGrid.RowCount := FileGrid.RowCount - 1
    else
    begin
      ListBox.Clear;
      FileGrid.Cells[0, 1] := '';
      FileGrid.Cells[1, 1] := '';
      FileGrid.Cells[2, 1] := '';
      FileGrid.Cells[3, 1] := '';
    end;
    ListBoxReload(Self);
    GridRenumerate(Self);
  end;
end;

procedure TSymLinkForm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  If ((Key = ord('c')) or (Key = ord('C'))) and (Shift = [ssCtrl]) then
    CopyButtonClick(Sender);
end;

procedure TSymLinkForm.mDeleteClick(Sender: TObject);
var
  Key: Word;
  Shift: TShiftState;
begin
  Key := 46; // del
  FileGridKeyDown(Sender, Key, Shift);
end;

procedure TSymLinkForm.mNewClick(Sender: TObject);
begin
  ListBox.Items.Clear;
  FileGrid.RowCount := 2;
  All_sizes := 0;
  All_Files := 0;
  FileGrid.Cells[0, 1] := '';
  FileGrid.Cells[1, 1] := '';
  FileGrid.Cells[2, 1] := '';
  FileGrid.Cells[3, 1] := '';
end;

procedure TSymLinkForm.mQuitClick(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TSymLinkForm.mCopyAllClick(Sender: TObject);
var
  buff: PChar;
  a: Integer;
  strs: Tstrings;
begin
  strs := TStringList.Create;
  try
    for a := 1 to FileGrid.RowCount - 1 do
      strs.Add(FileGrid.Cells[1, a] + FileGrid.Cells[2, a] + #09 + FileGrid.Cells[3, a]);
    buff := PChar(strs.Text);
    Clipboard.SetTextBuf(buff);
  finally
    strs.Free;
  end;
end;

procedure TSymLinkForm.mCopyNamesClick(Sender: TObject);
// only names to buffer
Var
  buff: PChar;
  a: Integer;
  strs: Tstrings;
begin
  strs := TStringList.Create;
  try
    for a := 1 to FileGrid.RowCount - 1 do
    begin
      strs.Add(FileGrid.Cells[2, a]);
    end;
    buff := PChar(strs.Text);
    Clipboard.SetTextBuf(buff);
  finally
    strs.Free;
  end;
end;

end.
