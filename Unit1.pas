unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  MyVisioUnit, Vcl.StdCtrls, Data.DB, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.VCLUI.Wait,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.ExtCtrls, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, FireDAC.Phys.SQLite, FireDAC.Phys.SQLiteDef,
  FireDAC.Stan.ExprFuncs, Vcl.Menus, StrUtils, Vcl.ComCtrls;

type
  TVST2SQL = class(TForm)
    FDConnection1: TFDConnection;
    FDQuery1: TFDQuery;
    DataSource1: TDataSource;
    VisioBox: TComboBox;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    OpenDatabase: TMenuItem;
    Exit: TMenuItem;
    About: TMenuItem;
    OnlyBox: TCheckBox;
    PageNo: TEdit;
    Label1: TLabel;
    FillFormat: TCheckBox;
    Label2: TLabel;
    LineFormat: TCheckBox;
    Character: TCheckBox;
    Scratch: TCheckBox;
    ShapeTransform: TCheckBox;
    Textfield: TCheckBox;
    UserDefined: TCheckBox;
    btnCommit: TButton;
    OpenDialog1: TOpenDialog;
    ProgressBar1: TProgressBar;
    Label3: TLabel;
    CloseDatabase1: TMenuItem;
    procedure OnClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCommitClick(Sender: TObject);
    procedure OpenDatabaseClick(Sender: TObject);
    procedure CheckSections();
    procedure CommitPage(FDQuery: TFDQuery; iPageNo: Integer; iShapeID: Integer; iSection: Integer; iColumn: Integer; iLine: Integer);
    procedure ExitClick(Sender: TObject);
    procedure AboutClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CloseDatabase1Click(Sender: TObject);

    private
    { Private declarations }
  public
    { Public declarations }
  end;

  {
   Projet : Learn Visio!
   Auteur : Sami N.
   Status : Finished.
   }

var
  VST2SQL: TVST2SQL;
  iDocID : integer;
  Connected: boolean=false;
  SelectedDoc: String;

  CharacterSection: boolean;
  FillFormatSection: boolean;
  LineFormatSection: boolean;
  ShapeTransformSection: boolean;
  ScratchSection: boolean;
  TextfieldSection: boolean;
  UserDefinedSection: boolean;
  OnlyBoxCheck: boolean;

  TableStructure: string = '(iFileID, iPageNo, iShapeID, iRow, sRowName, iSection, iColumn, sFormula, sNote)';

  filepath: string;
  TemplateTable: String;
  ShapesInfoTable: String;
  NameField: String;
  PathField: String;
  FileIDField: String;
  iFileID: Integer;


  i, j, iID, iLine, iPosition, Ok,  iPageNo, iPage_ID, iRowNameCount, iShapeCount, iShapeID  : integer;
  s, vsoFileName, sLabel, sType, sValue, sPropName, sPrompt, sName, sNameU, sFormula, sFullPathName, sRowName : string;

implementation

{$R *.dfm}
procedure TVST2SQL.CommitPage(FDQuery: TFDQuery; iPageNo: Integer; iShapeID: Integer; iSection: Integer; iColumn: Integer; iLine: Integer);
var
    query: string;
begin
     VisioGetStrCellFormula(iDocID, iPageNo, iShapeID,  iSection, iLine, iColumn, sFormula);
     VisioGetStrRowName(iDocID, iPageNo, iShapeID, visSectionProp, iLine, sName, sNameU);


     if sFormula = '' then begin
     query := 'INSERT INTO ' + ShapesInfoTable + TableStructure + ' VALUES(' + IntToStr(iFileID) + ',' + IntToStr(iPageNo) + ',' + IntToStr(iShapeID) + ',' + IntToStr(iLine) + ',"' + sName + '",' + IntToStr(iSection) + ','  + IntToStr(iColumn) + ', "", "" );';
     end
     else begin
        if ContainsText(sFormula, '"') then begin
         sFormula := StringReplace(sFormula, '"', '""', [rfReplaceAll, rfIgnoreCase]);
         query := 'INSERT INTO ' + ShapesInfoTable + TableStructure + ' VALUES(' +  IntToStr(iFileID) + ','+ IntToStr(iPageNo) + ',' + IntToStr(iShapeID) + ',' + IntToStr(iLine) + ',"' + sName + '",' + IntToStr(iSection) + ','  + IntToStr(iColumn) + ',"' + sFormula + '", "");';
        end
        else begin

         query := 'INSERT INTO ' + ShapesInfoTable + TableStructure + ' VALUES(' +  IntToStr(iFileID) + ','+ IntToStr(iPageNo) + ',' + IntToStr(iShapeID) + ',' + IntToStr(iLine) + ',"' + sName + '",' + IntToStr(iSection) + ','  + IntToStr(iColumn) + ',"' + sFormula + '", "");';
        end;

     end;

     FDQuery.ExecSQL(query);
     FDQuery.Close();

end;


procedure TVST2SQL.ExitClick(Sender: TObject);
begin
  VisioDiscardDoc(iDocID);
  VisioDisconnect;
  Connected := false;
  FDQuery1.Close();
  FDConnection1.Connected := false;
  Application.Terminate;
end;

procedure TVST2SQL.FormCreate(Sender: TObject);
begin
      FDConnection1.Connected := false;
end;

procedure TVST2SQL.CheckSections();
begin
     if Character.Checked then begin
        CharacterSection := true;
     end
     else begin
        CharacterSection := false;
     end;

     if FillFormat.Checked then begin
        FillFormatSection := true;
     end
     else begin
        FillFormatSection := false;
     end;

     if LineFormat.Checked then begin
        LineFormatSection := true;
     end
     else begin
        LineFormatSection := false;
     end;

     if ShapeTransform.Checked then begin
        ShapeTransformSection := true;
     end
     else begin
        ShapeTransformSection := false;
     end;

     if Scratch.Checked then begin
        ScratchSection := true;
     end
     else begin
        ScratchSection := false;
     end;

     if Textfield.Checked then begin
        TextfieldSection := true;
     end
     else begin
        TextfieldSection := false;
     end;

     if UserDefined.Checked then begin
        UserDefinedSection := true;
     end
     else begin
        UserDefinedSection := false;
     end;

     if OnlyBox.Checked then begin
        OnlyBoxCheck := true;
     end
     else begin
        OnlyBoxCheck := false;
     end;

end;







procedure TVST2SQL.CloseDatabase1Click(Sender: TObject);
begin
  VisioDiscardDoc(iDocID);
  VisioDisconnect;
  Connected := false;
  FDQuery1.Close();
  FDConnection1.Connected := false;
end;

procedure TVST2SQL.AboutClick(Sender: TObject);
begin
     ShowMessage('VST2SQL | Delphi 10.3 | Author : Sami N.');
end;

procedure TVST2SQL.btnCommitClick(Sender: TObject);
var
     query: String;
     debug: string;
     vsoFileName: String;

begin
    SelectedDoc := VisioBox.Text;
    query := 'SELECT ' + FileIDField + ',' + PathField + ' FROM ' + TemplateTable + ' WHERE ' + NameField + ' LIKE "%' + SelectedDoc + '%";';
    try
    FDQuery1.Open(query);
    iFileID := FDQuery1.Fields.FieldByName(FileIDField).AsInteger;
    sFullPathName := FDQuery1.Fields.FieldByName(PathField).AsString;
    except begin
       Application.MessageBox('Invalid or Incorrect Field Name.', 'Error', MB_OK or MB_ICONERROR);
    end;

    end;


    vsoFileName:= sFullPathName;
    VisioConnect();
    Try
       iDocID:=VisioStrOpenDoc(vsoFileName);
    except begin
       Application.MessageBox('Cannot open or access visio file.', 'Error', MB_OK or MB_ICONERROR);
    end;

    End;

    iPageNo := StrToInt(PageNo.Text);
    iPage_ID:=VisioGetPageSheetID(iDocID, iPageNo);
    iShapeCount :=  VisioGetShapeCount(iDocID,  iPageNo, False);
    s:=inttostr(iShapeCount);
    Connected := true;
    CheckSections();

    if OnlyBoxCheck then begin
       iRowNameCount := VisioGetRowCount(iDocID, iPageNo, iPage_ID, visSectionProp);
       for j := 0 to iRowNameCount-1 do begin
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsLabel, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsPrompt, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsSortKey, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsType, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsFormat, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsValue, j);
       end;
    end
    else begin

           //Page ShapeData
           iRowNameCount := VisioGetRowCount(iDocID, iPageNo, iPage_ID, visSectionProp);
            for j := 0 to iRowNameCount-1 do begin
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsLabel, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsPrompt, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsSortKey, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsType, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsFormat, j);
            CommitPage(FDQuery1, iPageNo, iShapeID, visSectionProp, visCustPropsValue, j);
           end;
           // Other sections
           if iShapeCount >= 0 then begin
              ProgressBar1.Max := iShapeCount-1;
              ProgressBar1.Min := 0;
           end;

           
           for j := 0 to iShapeCount-1 do begin

              if CharacterSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
              iRowNameCount := VisioGetRowCount(iDocID, iPageNo, iShapeID, visSectionCharacter);
              for i := 0 to iRowNameCount-1 do begin
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionCharacter, visCharacterSize, i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionCharacter, visCharacterColor, i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionCharacter, visCharacterFont, i);
              end;

              end;


              if FillFormatSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowFill, visFillBkgnd);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowFill, visFillPattern );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowFill, visFillForegnd );
              end;

              if LineFormatSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowLine , visLineColor );
              end;

              if ShapeTransformSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormPinX);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormPinY );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormWidth );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormHeight );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormLocPinX );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormLocPinY );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionObject , visRowXFormOut , visXFormAngle );

              end;

              if ScratchSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
              iRowNameCount := VisioGetRowCount(iDocID, iPageNo, iShapeID, visSectionScratch);
              for i := 0 to iRowNameCount-1 do begin
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionScratch, visScratchX , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionScratch, visScratchY , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionScratch, visScratchA , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionScratch, visScratchB , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionScratch, visScratchC , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionScratch, visScratchD , i);
              end;

              end;

              if TextfieldSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
              iRowNameCount := VisioGetRowCount(iDocID, iPageNo, iShapeID, visSectionTextField);
              for i := 0 to iRowNameCount-1 do begin
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionTextField, visFieldType  , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionTextField, visFieldFormat  , i);
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionTextField, visFieldCell  , i);
              end;

              end;

              if UserDefinedSection then begin
              iShapeID := VisioGetShapeID(iDocID, iPageNo, j, false);
              iRowNameCount := VisioGetRowCount(iDocID, iPageNo, iShapeID, visSectionUser);
              for i := 0 to iRowNameCount-1 do begin
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionUser, visFieldType  , visUserValue );
                  CommitPage(FDQuery1, iPageNo, iShapeID, visSectionUser, visFieldFormat  , visUserValue );
              end;

              end;

              ProgressBar1.Position := j;
              Label3.Caption := IntToStr(j) + ' / ' + IntToStr(iShapeCount);
           end;

           Label3.Caption := Label3.Caption + ' Done.';

    end;

end;





procedure TVST2SQL.OnClose(Sender: TObject; var Action: TCloseAction);
begin
  VisioDiscardDoc(iDocID);
  VisioDisconnect;
  Connected := false;
  FDQuery1.Close();
  FDConnection1.Connected := false;
  Application.Terminate;
end;




procedure TVST2SQL.OpenDatabaseClick(Sender: TObject);
var
  openquery: string;

begin
      if FDConnection1.Connected then begin
      Application.MessageBox('Already connected to another database, please close the current one before opening a new one.', 'Error', MB_OK or MB_ICONERROR);
      end
      else begin
          OpenDialog1.Filter := 'Database (*.db*)|*.db*';
          if OpenDialog1.Execute then begin
        filepath := OpenDialog1.filename;
        end;
        FDConnection1.ConnectionString := 'DriverID=SQLite;Database=' + filepath + ';';
        FDConnection1.Connected := true;
        FDConnection1.Open;

        VST2SQL.Caption  := 'VST2SQL [Connected]';

        TemplateTable    := InputBox('SQLite Query','Enter the template table:', 'tbl_Templates');

        NameField        := InputBox('SQLite Query','Enter the filename field:', 'sName');

        PathField        := InputBox('SQLite Query','Enter the filepath field:', 'sFullPathName');

        FileIDField      := InputBox('SQLite Query','Enter the FileID field:', 'iFileID');

        ShapesInfoTable  := InputBox('SQLite Query','Enter the ShapeInfo table:', 'visio');

        if (TemplateTable = '') or (Namefield = '') or (Pathfield = '') or (FileIDfield = '') or (ShapesInfoTable = '') then begin
            ShowMessage('Invalid Input');
        end;

        openquery := 'SELECT * FROM ' + TemplateTable + ';';
        FDQuery1.Open(openquery);

        for i := 0 to VisioBox.Items.Count do  begin
          VisioBox.Items.Delete(0);
        end;


        while not FDQuery1.Eof do begin
          VisioBox.Items.Add(FDQuery1.Fields.FieldByName(NameField).AsString);
          FDQuery1.Next;
        end;
        FDQuery1.Close;
      end;

end;

end.


