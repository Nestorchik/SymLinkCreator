unit SymLinkUnit;

interface

uses
  Windows, SysUtils, Controls, Forms, Messages, Dialogs, Grids, Clipbrd, ExtCtrls, Menus, StdCtrls, Classes, Graphics,
  Vcl.ComCtrls, IdBaseComponent, IdMessage, Vcl.TabNotBk, DragDrop, DropSource, DragDropFile, DropTarget, DragDropFormats,
  ImgList, ActnList, Actions, ImageList, Types, example, Vcl.Buttons, Vcl.ToolWin, Vcl.OleCtrls, SHDocVw;

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
    FileGrid: TStringGrid;
    mDelete: TMenuItem;
    mmDelClick: TMenuItem;
    msgTimer: TTimer;
    BottomGridPanel: TGridPanel;
    AskDirs: TCheckBox;
    CopySizes: TCheckBox;
    CopyPaths: TCheckBox;
    DoDirs: TCheckBox;
    ShellExecute: TBitBtn;
    CopyButton: TBitBtn;
    ListBox: TListBox;
    BlinkTimer: TTimer;
    mOpen: TMenuItem;
    StatusBar: TStatusBar;
    WebBrowser: TWebBrowser;
    procedure FormCreate(Sender: TObject);
    procedure mSaveClick(Sender: TObject);
    procedure mSaveAllToTExtClick(Sender: TObject);
    procedure mNewClick(Sender: TObject);
    procedure mQuitClick(Sender: TObject);
    procedure FileGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mCopyAllClick(Sender: TObject);
    procedure mCopyNamesClick(Sender: TObject);
    procedure CopyButtonClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mmOpenClick(Sender: TObject);
    procedure LeftPanelMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure mDeleteClick(Sender: TObject);
    procedure ShellExecuteClick(Sender: TObject);
    procedure msgTimerTimer(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BlinkTimerTimer(Sender: TObject);
    procedure ListBoxReload(Sender: TObject);
    procedure GridRenumerate(Sender: TObject);
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
  sNeedFiles, smFile, smFileHint, smOpen, smOpenHint, smSave, smSaveHint, smNew, smNewHint, smQuit, smQuitHint, smFileList, smFileListHint, smDelete, smDeleteHint, smClear, smClearHint, smCopyAll, smCopyAllHint, smCopyNAmes, smCopyNAmesHint, smpNew,
    smpOpen, smpSave, smpCopyAll, smpCopyNames, smpDelete, smpQuit: string;

implementation

uses
  ShellApi, IniFiles;

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

procedure TSymLinkForm.FormCreate(Sender: TObject);
var
  i, a: Integer;
  Ini, langIni: TIniFile;
  List: TStringList;
  f: TextFile;
  s: String;
begin
  // accept drag files to program
  DragAcceptFiles(SymLinkForm.Handle, true);
  // timeStamps
  longTimeStamp := FormatDateTime('yyyymmdd-hhnnsszzz', now);
  hourTimeStamp := FormatDateTime('yyyymmdd-hh', now);
  // load ini-files
  Ini := TIniFile.Create(Extractfilepath(paramstr(0)) + 'SymLinkCreator.ini');
  SymLinkForm.Width := Ini.ReadInteger('FormSize', 'Width', 840);
  SymLinkForm.Height := Ini.ReadInteger('FormSize', 'Height', 700);
  SymLinkForm.Left := Ini.ReadInteger('FormPosition', 'X', 100);
  SymLinkForm.Top := Ini.ReadInteger('FormPosition', 'Y', 100);
  DoDirs.Checked := Ini.ReadBool('DoDirs', 'Checked', false);
  AskDirs.Checked := Ini.ReadBool('AskDirs', 'Checked', false);
  CopySizes.Checked := Ini.ReadBool('CopySizes', 'Checked', false);
  CopyPaths.Checked := Ini.ReadBool('CopyPaths', 'Checked', false);
  SymLinkForm.FileGrid.ColWidths[1] := Ini.ReadInteger('FileGridColWidth', 'Width1', 370);
  SymLinkForm.FileGrid.ColWidths[2] := Ini.ReadInteger('FileGridColWidth', 'Width2', 310);
  SymLinkForm.FileGrid.ColWidths[3] := Ini.ReadInteger('FileGridColWidth', 'Width3', 85);
  MaxDragFiles := Ini.ReadInteger('MaxDragFiles', 'Files', 25);
  sHelpFile := Ini.ReadString('Help', 'File', 'SimLinkCreatorHelpEn.html');
  TabNotebook.ActivePage := Ini.ReadString('Pages', 'ActivePAge', 'Help');
  logFilePath := Ini.ReadString('Logs', 'fPath', Extractfilepath(paramstr(0)) + 'Logs');
  logFileName := Ini.ReadString('Logs', 'fName', 'SimLinkCreator');
  logFileExt := Ini.ReadString('Logs', 'fExt', '.log');
  batFilePath := Ini.ReadString('Bat', 'fPath', Extractfilepath(paramstr(0)) + 'Bat');
  batFileName := Ini.ReadString('Bat', 'fName', 'SimLinkCreator');
  batFileExt := Ini.ReadString('Bat', 'fExt', '.bat');
  slastFileName := Ini.ReadString('LastFile', 'File', Extractfilepath(paramstr(0)) + 'SymLinkCreator_last.txt');
  Ini.Free;
  // open lang files
  langIni := TIniFile.Create(Extractfilepath(paramstr(0)) + 'SymLinkCreator_lang.ini');
  sNeedFiles := langIni.ReadString('ENG', 'NeedFiles', 'Drag files/folders then start "Create links"');
  FolderStr := langIni.ReadString('ENG', 'FolderStr', 'Folder');
  NotAFolder := langIni.ReadString('ENG', 'NotAFolder', 'First string must be FOLDER!!!');
  srtUnavail := langIni.ReadString('ENG', 'Unavailable', 'Unavailable');
  sPath := langIni.ReadString('ENG', 'Path', 'Path');
  sFile := langIni.ReadString('ENG', 'File', 'File');
  sSize := langIni.ReadString('ENG', 'Sise', 'Size');
  sFiles := langIni.ReadString('ENG', 'TabFiles', 'Files');
  sList := langIni.ReadString('ENG', 'TabList', 'Flat list');
  sHelp := langIni.ReadString('ENG', 'TabHelp', 'Help');
  sMsgInclideFiles := langIni.ReadString('ENG', 'sMsgInclideFiles', 'Include files in them instead of folders?');
  sMsgDlgCaption := langIni.ReadString('ENG', 'sMsgDlgCaption', 'The logic of enabling folders');
  srtUnavail := langIni.ReadString('ENG', 'srtUnavail', 'Unavailable');
  eFromEncode := langIni.ReadString('ENG', 'eFromEncode', '1251');
  eToEncode := langIni.ReadString('ENG', 'eToEncode', '866');
  sMaxDragFiles := langIni.ReadString('ENG', 'sMaxDragFiles', ' ' + IntToStr(MaxDragFiles) + ' files max - ini-file settings');
  sNoData := langIni.ReadString('ENG', 'NoData', 'No data');
  sDoDirs := langIni.ReadString('ENG', 'DoDirs', 'files instead of folders or >');
  sDoDirsHint := langIni.ReadString('ENG', 'DoDirsHint', 'Include files in folders instead of folders in the list');
  sDoDirs := langIni.ReadString('ENG', 'DoDirs', 'files instead of folders or >');
  sAskDirs := langIni.ReadString('ENG', 'AskDirs', 'ask all time');
  sAskDirsHint := langIni.ReadString('ENG', 'AskDirsHint', 'Ask everyone about the folders in the list');
  sCopyButton := langIni.ReadString('ENG', 'CopyButton', 'Copy file list');
  sCopyButtonHint := langIni.ReadString('ENG', 'CopyButtonHint', 'Copy file list to buffer');
  sCopySizes := langIni.ReadString('ENG', 'CopySizes', '+ size');
  sCopySizesHint := langIni.ReadString('ENG', 'sCopySizesHint', 'Add the file size to the file names as well');
  sCopyPaths := langIni.ReadString('ENG', 'CopyPaths', '+ path');
  sCopyPathsHint := langIni.ReadString('ENG', 'sCopyPathsHint', 'Add the full paths to the files');
  sShellExecute := langIni.ReadString('ENG', 'ShellExecute', 'Create Links');
  sShellExecuteHint := langIni.ReadString('ENG', 'ShellExecuteHint', 'Run script and create symbolic links');
  // load all menu lang strings
  smFile := langIni.ReadString('ENG-Menu', smFile, 'File');
  smFileHint := langIni.ReadString('ENG-Menu', smFileHint, 'File operations');
  smOpen := langIni.ReadString('ENG-Menu', smOpen, 'Open');
  smOpenHint := langIni.ReadString('ENG-Menu', smOpenHint, 'Open saved file list');
  smSave := langIni.ReadString('ENG-Menu', smSave, 'Save');
  smSaveHint := langIni.ReadString('ENG-Menu', smSaveHint, 'Save file list');
  smNew := langIni.ReadString('ENG-Menu', smNew, 'New');
  smNewHint := langIni.ReadString('ENG-Menu', smNewHint, 'New file list');
  smQuit := langIni.ReadString('ENG-Menu', smQuit, 'Quit');
  smQuitHint := langIni.ReadString('ENG-Menu', smQuitHint, 'Finish and exit');
  smFileList := langIni.ReadString('ENG-Menu', smFileList, 'File list');
  smFileListHint := langIni.ReadString('ENG-Menu', smFileListHint, 'Operations with filelist');
  smDelete := langIni.ReadString('ENG-Menu', smDelete, 'Delete');
  smDeleteHint := langIni.ReadString('ENG-Menu', smDeleteHint, 'Delete current string');
  smClear := langIni.ReadString('ENG-Menu', smClear, 'Clear');
  smClearHint := langIni.ReadString('ENG-Menu', smClearHint, 'Clear all');
  smCopyAll := langIni.ReadString('ENG-Menu', smCopyAll, 'Copy all');
  smCopyAllHint := langIni.ReadString('ENG-Menu', smCopyAllHint, 'Copy all to buffer');
  smCopyNAmes := langIni.ReadString('ENG-Menu', smCopyNAmes, 'Copy only names');
  smCopyNAmesHint := langIni.ReadString('ENG-Menu', smCopyNAmesHint, 'Copy only names to buffer');
  smpNew := langIni.ReadString('ENG-Menu', smpNew, 'New');
  smpOpen := langIni.ReadString('ENG-Menu', smpOpen, 'Open');
  smpSave := langIni.ReadString('ENG-Menu', smpSave, 'Save');
  smpCopyAll := langIni.ReadString('ENG-Menu', smpCopyAll, 'Copy all');
  smpCopyNames := langIni.ReadString('ENG-Menu', smpCopyNames, 'Copy names');
  smpDelete := langIni.ReadString('ENG-Menu', smpDelete, 'Delete');
  smpQuit := langIni.ReadString('ENG-Menu', smpQuit, 'Quit');
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
  MainMenu.Items.Items[1].Items[3].Caption := smCopyNAmes;
  MainMenu.Items.Items[1].Items[3].Hint := smCopyNAmesHint;
  // init pop-up menu lang strings
  pMenu.Items[0].Caption := smpNew;
  pMenu.Items[1].Caption := smpOpen;
  pMenu.Items[2].Caption := smpSave;
  pMenu.Items[3].Caption := smpCopyAll;
  pMenu.Items[4].Caption := smpCopyNames;
  pMenu.Items[5].Caption := smpDelete;
  pMenu.Items[6].Caption := smpQuit;
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
  // try load Help page without exceptions
  try
    if FileExists(sHelpFile) then
      WebBrowser.Navigate(Extractfilepath(paramstr(0)) + sHelpFile)
    else
      TabNotebook.Pages.Delete(2);
  finally
  end;
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
end;

procedure TSymLinkForm.FormDestroy(Sender: TObject);
var
  Ini: TIniFile;
  langIni: TIniFile;
  StringList: TStringList;
  i, a: Integer;
  f: TextFile;
  str: string;
begin
  // Save INI-file
  Ini := TIniFile.Create(Extractfilepath(paramstr(0)) + 'SymLinkCreator.ini');
  Ini.WriteInteger('FormSize', 'Width', SymLinkForm.Width);
  Ini.WriteInteger('FormSize', 'Height', SymLinkForm.Height);
  Ini.WriteInteger('FormPosition', 'X', SymLinkForm.Left);
  Ini.WriteInteger('FormPosition', 'Y', SymLinkForm.Top);
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
  Ini.Free;
  // Save main lang strings
  langIni := TIniFile.Create(Extractfilepath(paramstr(0)) + 'SymLinkCreator_lang.ini');
  langIni.WriteString('ENG', 'FilderStr', FolderStr);
  langIni.WriteString('ENG', 'NeedFiles', sNeedFiles);
  langIni.WriteString('ENG', 'folderFirts', NotAFolder);
  langIni.WriteString('ENG', 'Unavailable', srtUnavail);
  langIni.WriteString('ENG', 'Path', sPath);
  langIni.WriteString('ENG', 'File', sFile);
  langIni.WriteString('ENG', 'Sise', sSize);
  langIni.WriteString('ENG', 'TabFiles', sFiles);
  langIni.WriteString('ENG', 'TabList', sList);
  langIni.WriteString('ENG', 'TabHelp', sHelp);
  langIni.WriteString('ENG', 'sMsgInclideFiles', sMsgInclideFiles);
  langIni.WriteString('ENG', 'sMsgDlgCaption', sMsgDlgCaption);
  langIni.WriteString('ENG', 'srtUnavail', srtUnavail);
  langIni.WriteString('ENG', 'eFromEncode', eFromEncode);
  langIni.WriteString('ENG', 'eToEncode', eToEncode);
  langIni.WriteString('ENG', 'sMaxDragFiles', sMaxDragFiles);
  langIni.WriteString('ENG', 'NoData', sNoData);
  langIni.WriteString('ENG', 'DoDirs', sDoDirs);
  langIni.WriteString('ENG', 'DoDirsHint', sDoDirsHint);
  langIni.WriteString('ENG', 'CopyButton', sCopyButton);
  langIni.WriteString('ENG', 'CopyButtonHint', sCopyButtonHint);
  langIni.WriteString('ENG', 'CopySizes', sCopySizes);
  langIni.WriteString('ENG', 'CopySizesHint', sCopySizesHint);
  langIni.WriteString('ENG', 'CopyPaths', sCopyPaths);
  langIni.WriteString('ENG', 'CopyPathsHint', sCopyPathsHint);
  langIni.WriteString('ENG', 'ShellExecute', sShellExecute);
  langIni.WriteString('ENG', 'ShellExecuteHint', sShellExecuteHint);
  // Save menu lang strings
  langIni.WriteString('ENG-Menu', 'smFile', smFile);
  langIni.WriteString('ENG-Menu', 'smFileHint', smFileHint);
  langIni.WriteString('ENG-Menu', 'smOpen', smOpen);
  langIni.WriteString('ENG-Menu', 'smOpenHint', smOpenHint);
  langIni.WriteString('ENG-Menu', 'smSave', smSave);
  langIni.WriteString('ENG-Menu', 'smSaveHint', smSaveHint);
  langIni.WriteString('ENG-Menu', 'smNew', smNew);
  langIni.WriteString('ENG-Menu', 'smNewHint', smNewHint);
  langIni.WriteString('ENG-Menu', 'smQuit', smQuit);
  langIni.WriteString('ENG-Menu', 'smQuitHint', smQuitHint);
  langIni.WriteString('ENG-Menu', 'smFileList', smFileList);
  langIni.WriteString('ENG-Menu', 'smFileListHint', smFileListHint);
  langIni.WriteString('ENG-Menu', 'smDelete', smDelete);
  langIni.WriteString('ENG-Menu', 'smDeleteHint', smDeleteHint);
  langIni.WriteString('ENG-Menu', 'smClear', smClear);
  langIni.WriteString('ENG-Menu', 'smClearHint', smClearHint);
  langIni.WriteString('ENG-Menu', 'smCopyAll', smCopyAll);
  langIni.WriteString('ENG-Menu', 'smCopyAllHint', smCopyAllHint);
  langIni.WriteString('ENG-Menu', 'smCopyNAmes', smCopyNAmes);
  langIni.WriteString('ENG-Menu', 'smCopyNAmesHint', smCopyNAmesHint);
  langIni.WriteString('ENG-Menu', 'smpNew', smpNew);
  langIni.WriteString('ENG-Menu', 'smpOpen', smpOpen);
  langIni.WriteString('ENG-Menu', 'smpSave', smpSave);
  langIni.WriteString('ENG-Menu', 'smpCopyAll', smpCopyAll);
  langIni.WriteString('ENG-Menu', 'smpCopyNAmes', smpCopyNames);
  langIni.WriteString('ENG-Menu', 'smpDelete', smpDelete);
  langIni.WriteString('ENG-Menu', 'smpQuit', smpQuit);
  langIni.Free;
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

procedure findlastmath(str1, str2: string; var findtopos: Integer);
var
  a: Integer;
begin
  if length(str1) < findtopos then
    findtopos := length(str1);
  if length(str2) < findtopos then
    findtopos := length(str2);
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
    application.ProcessMessages;
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
  st: string;
begin
  If AskDirs.Checked then // How many files are dropped
  begin
    DoDirs.Checked := application.MessageBox(PChar(sMsgInclideFiles), PChar(sMsgDlgCaption), MB_YESNO) = IDYes;
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
  s, s1, line1: String;
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

procedure TSymLinkForm.mSaveAllToTExtClick(Sender: TObject);
Var
  fs: Tfilestream;
  st: PChar;
  a: Integer;
begin
  if SaveFileDialog.FileName = '' then
    SaveFileDialog.FileName := '*.txt';
  If SaveFileDialog.Execute then
  begin
  end;
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
    application.ProcessMessages;
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

  if strFolder.length = 0 then
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

  {
    if ((strFolder = FolderStr) and (FileGrid.Cells[1, 2] <> '')) then
    begin
    ShellExecute.Enabled := true;
    StatusBar.Panels.Items[0].Text := 'OK';
    end
    else
    begin
    ShellExecute.Enabled := false;
    StatusBar.SimpleText := '123';
    end;

    if ((strFolder.length > 0) and (strFolder <> FolderStr)) and (FileGrid.Cells[2, 2] <> '') then
    begin
    StatusBar.Panels.Items[0].Text := NotAFolder;
    ShellExecute.Enabled := false;
    mNewClick(Self);
    end;
  }
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

procedure TSymLinkForm.BlinkTimerTimer(Sender: TObject);
begin
  {
    If you've read everything here, then accept my fervent greetings! ))))))))
    Do you want me to tell you a nursery rhyme?)))))
  }
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

procedure TSymLinkForm.LeftPanelMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  fileContent, FileName: string; { NOT USE - TEMP PLUG }
begin
  if (Button = mbLeft) and DragDetectPlus(Handle, Point(X, Y)) then
  begin
    FileName := '123.txt';
    fileContent := '123-123-123';
    // DropEmptySource1.Execute;
  end;
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
  application.Terminate;
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
{ only names to buf }
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
