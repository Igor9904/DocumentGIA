unit GIADoc;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, WordXP, OleServer, Win.ComObj, WinApi.ActiveX,
  Vcl.ComCtrls, Vcl.Imaging.pngimage, inifiles, Vcl.Buttons;

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    P_Vypiska: TPanel;
    Label1: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    M_FIOVyp: TMemo;
    E_NumGrup: TEdit;
    B_Vypiska: TButton;
    M_Diplom: TMemo;
    E_ProcVyp: TEdit;
    P_Titul: TPanel;
    Label3: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    M_FIOTitul: TMemo;
    M_NumTel: TMemo;
    E_NumGrupTitul: TEdit;
    B_Titul: TButton;
    E_ProcTitul: TEdit;
    TabSheet3: TTabSheet;
    ScrollBox1: TScrollBox;
    Panel1: TPanel;
    Image1: TImage;
    Memo1: TMemo;
    Panel2: TPanel;
    Label16: TLabel;
    E_NumProt: TEdit;
    Label17: TLabel;
    Panel3: TPanel;
    Label18: TLabel;
    Label19: TLabel;
    CB_StudFIO: TComboBox;
    CheckBox1: TCheckBox;
    E_StudFIOManual: TEdit;
    Panel4: TPanel;
    Label20: TLabel;
    E_Grup: TEdit;
    Label21: TLabel;
    E_Spec: TEdit;
    Label22: TLabel;
    Label23: TLabel;
    E_PredsedatelFIO: TEdit;
    Label24: TLabel;
    Panel6: TPanel;
    Label25: TLabel;
    GridPanel1: TGridPanel;
    Panel5: TPanel;
    GEK1: TMemo;
    Panel7: TPanel;
    GEK2: TMemo;
    Panel8: TPanel;
    GEK3: TMemo;
    Panel9: TPanel;
    GEK4: TMemo;
    Panel10: TPanel;
    GEK5: TMemo;
    Panel11: TPanel;
    GEK6: TMemo;
    Panel12: TPanel;
    GEK7: TMemo;
    Panel13: TPanel;
    GEK8: TMemo;
    CB_GEK1: TCheckBox;
    CB_GEK2: TCheckBox;
    CB_GEK3: TCheckBox;
    CB_GEK4: TCheckBox;
    CB_GEK5: TCheckBox;
    CB_GEK6: TCheckBox;
    CB_GEK7: TCheckBox;
    CB_GEK8: TCheckBox;
    Button1: TButton;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Panel14: TPanel;
    E_Day1: TEdit;
    Label29: TLabel;
    Label30: TLabel;
    Panel15: TPanel;
    Panel16: TPanel;
    Label31: TLabel;
    EMP: TComboBox;
    Label32: TLabel;
    Panel17: TPanel;
    Label33: TLabel;
    Label34: TLabel;
    Stacionar: TComboBox;
    Panel18: TPanel;
    Label35: TLabel;
    Label36: TLabel;
    Poliklinika: TComboBox;
    Panel19: TPanel;
    Label37: TLabel;
    Label38: TLabel;
    SLR: TComboBox;
    Panel20: TPanel;
    Label39: TLabel;
    Itog1: TComboBox;
    Label40: TLabel;
    E_Day2: TEdit;
    Label41: TLabel;
    Label42: TLabel;
    Panel21: TPanel;
    Label43: TLabel;
    E_TestCor: TEdit;
    Panel22: TPanel;
    Label44: TLabel;
    Itog2: TComboBox;
    Button2: TButton;
    Button3: TButton;
    Label45: TLabel;
    E_Day3: TEdit;
    Label46: TLabel;
    Panel23: TPanel;
    Label47: TLabel;
    E_NumByl: TEdit;
    GridPanel2: TGridPanel;
    Panel24: TPanel;
    Great: TMemo;
    CB_AnsOtl: TCheckBox;
    Panel25: TPanel;
    Good: TMemo;
    CB_AnsGood: TCheckBox;
    Panel26: TPanel;
    Passable: TMemo;
    CB_AnsUd: TCheckBox;
    Panel27: TPanel;
    Substandard: TMemo;
    CB_Ansneud: TCheckBox;
    Button4: TButton;
    Label48: TLabel;
    Label49: TLabel;
    M_DopVop: TMemo;
    Label50: TLabel;
    GridPanel3: TGridPanel;
    Panel28: TPanel;
    Dop_Great: TMemo;
    CB_DopAnsOtl: TCheckBox;
    Panel29: TPanel;
    Dop_Good: TMemo;
    CB_DopAnsGood: TCheckBox;
    Panel30: TPanel;
    Dop_Passable: TMemo;
    CB_DopAnsUd: TCheckBox;
    Panel31: TPanel;
    Dop_Substandard: TMemo;
    CB_DopAnsNeud: TCheckBox;
    Button5: TButton;
    Panel32: TPanel;
    Label51: TLabel;
    Itog3: TComboBox;
    Button6: TButton;
    Panel33: TPanel;
    Label52: TLabel;
    ItogGIA: TComboBox;
    Button7: TButton;
    Label53: TLabel;
    M_Comm: TMemo;
    Panel35: TPanel;
    Label54: TLabel;
    Label55: TLabel;
    E_Predsedatel: TEdit;
    Label322: TLabel;
    E_Secretar: TEdit;
    Button8: TButton;
    E_ProcProt: TEdit;
    E_Month3: TEdit;
    SpeedButton1: TSpeedButton;
    Button9: TButton;
    procedure B_TitulClick(Sender: TObject);
    procedure B_VypiskaClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ScrollBox1MouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure ScrollBox1MouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure CB_GEK1Click(Sender: TObject);
    procedure CB_GEK2Click(Sender: TObject);
    procedure CB_GEK3Click(Sender: TObject);
    procedure CB_GEK4Click(Sender: TObject);
    procedure CB_GEK5Click(Sender: TObject);
    procedure CB_GEK6Click(Sender: TObject);
    procedure CB_GEK7Click(Sender: TObject);
    procedure CB_GEK8Click(Sender: TObject);
    procedure CB_AnsOtlClick(Sender: TObject);
    procedure CB_AnsGoodClick(Sender: TObject);
    procedure CB_AnsUdClick(Sender: TObject);
    procedure CB_AnsneudClick(Sender: TObject);
    procedure CB_DopAnsOtlClick(Sender: TObject);
    procedure CB_DopAnsGoodClick(Sender: TObject);
    procedure CB_DopAnsUdClick(Sender: TObject);
    procedure CB_DopAnsNeudClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  komisiya: integer = 0;
  character: integer = 0;
  dopchar: integer = 0;
const
  PREFORM = '\preform';
implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  CB_GEK1.Checked:= False;
  CB_GEK2.Checked:= False;
  CB_GEK3.Checked:= False;
  CB_GEK4.Checked:= False;
  CB_GEK5.Checked:= False;
  CB_GEK6.Checked:= False;
  CB_GEK7.Checked:= False;
  CB_GEK8.Checked:= False;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  EMP.ItemIndex:= -1;
  Stacionar.ItemIndex:= -1;
  Poliklinika.ItemIndex:= -1;
  SLR.ItemIndex:= -1;
  Itog1.ItemIndex:= -1;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  E_TestCor.Text:= '';
  Itog2.ItemIndex:= -1;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  CB_AnsOtl.Checked:= False;
  CB_AnsGood.Checked:= False;
  CB_AnsUd.Checked:= False;
  CB_Ansneud.Checked:= False;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  CB_DopAnsOtl.Checked:= False;
  CB_DopAnsGood.Checked:= False;
  CB_DopAnsUd.Checked:= False;
  CB_DopAnsNeud.Checked:= False;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
  Itog3.ItemIndex:= -1;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
  ItogGIA.ItemIndex:= -1;
end;

procedure TForm1.Button8Click(Sender: TObject);
var
  TempleateFileName: string;
  WordApp, Document: OLEVariant;
  NameDOT, tmp_str, s: string;
  BookmarkName: string;
  Range, varcol: OLEVariant;
  i,j: integer;

  function CompareBm(ABmName: string; const AName: string): boolean;
  // проверяет наличие закладок
  var
    i: integer;
  begin
    i := Pos('__', ABmName);
    If i > 0 then
      Delete(ABmName, i, Length(ABmName) - i + 1);
    Result := SameText(ABmName, AName);
  end;

begin
    NameDOT := '\Protokol.dot';
    E_ProcTitul.Text:= '';
    if not CheckBox1.Checked then
      E_ProcProt.Text:= 'Создаю: '+CB_StudFIO.Text
    else
      E_ProcProt.Text:= 'Создаю: '+E_StudFIOManual.Text;

    If NameDOT <> '' then begin
          TempleateFileName := ExtractFileDir(Application.ExeName) + PathDelim + 'Templates' + NameDOT; // шаблон
          try
            WordApp := CreateOleObject('Word.Application');
          except
            On E: Exception do
            begin
              MessageBox(Application.Handle,
                PChar('Не удалось запустить MS Word!'#13#10 + E.Message), 'Ошибка',
                MB_OK + MB_ICONERROR + MB_SYSTEMMODAL);
              Exit;
            end;
          end;
          Document := WordApp.Documents.add(TempleateFileName);

          for i := Document.Bookmarks.Count downto 1 do begin // начинает перебирать все закладки в шаблоне
            BookmarkName := Document.Bookmarks.Item(i).Name;
            Range := Document.Bookmarks.Item(i).Range;
            If CompareBm(BookmarkName, 'НомерПрот') then begin
              if Trim(E_NumProt.Text) <> '' then
                Range.Text:= E_NumProt.Text;
            end;

            If CompareBm(BookmarkName, 'ФИО') then begin
              if not CheckBox1.Checked then
                Range.Text:= CB_StudFIO.Text
              else
                Range.Text:= E_StudFIOManual.Text;
            end;

            if CompareBm(BookmarkName, 'Группа') then
              Range.Text:= E_Grup.Text;

            if CompareBm(BookmarkName, 'Спец') then
              Range.Text:= E_Spec.Text;

            if CompareBm(BookmarkName, 'Председатель_ФИО') then
              Range.Text:= E_PredsedatelFIO.Text;

            If CompareBm(BookmarkName, 'Член1') then begin
              case komisiya of
              1: Range.Text:= GEK1.Lines[0];
              2: Range.Text:= GEK2.Lines[0];
              3: Range.Text:= GEK3.Lines[0];
              4: Range.Text:= GEK4.Lines[0];
              5: Range.Text:= GEK5.Lines[0];
              6: Range.Text:= GEK6.Lines[0];
              7: Range.Text:= GEK7.Lines[0];
              8: Range.Text:= GEK8.Lines[0];
              end;
            end;

            If CompareBm(BookmarkName, 'Член2') then begin
              case komisiya of
              1: Range.Text:= GEK1.Lines[1];
              2: Range.Text:= GEK2.Lines[1];
              3: Range.Text:= GEK3.Lines[1];
              4: Range.Text:= GEK4.Lines[1];
              5: Range.Text:= GEK5.Lines[1];
              6: Range.Text:= GEK6.Lines[1];
              7: Range.Text:= GEK7.Lines[1];
              8: Range.Text:= GEK8.Lines[1];
              end;
            end;

            If CompareBm(BookmarkName, 'Член3') then begin
              case komisiya of
              1: Range.Text:= GEK1.Lines[2];
              2: Range.Text:= GEK2.Lines[2];
              3: Range.Text:= GEK3.Lines[2];
              4: Range.Text:= GEK4.Lines[2];
              5: Range.Text:= GEK5.Lines[2];
              6: Range.Text:= GEK6.Lines[2];
              7: Range.Text:= GEK7.Lines[2];
              8: Range.Text:= GEK8.Lines[2];
              end;
            end;

            If CompareBm(BookmarkName, 'Член4') then begin
              case komisiya of
              1: Range.Text:= GEK1.Lines[3];
              2: Range.Text:= GEK2.Lines[3];
              3: Range.Text:= GEK3.Lines[3];
              4: Range.Text:= GEK4.Lines[3];
              5: Range.Text:= GEK5.Lines[3];
              6: Range.Text:= GEK6.Lines[3];
              7: Range.Text:= GEK7.Lines[3];
              8: Range.Text:= GEK8.Lines[3];
              end;
            end;

            If CompareBm(BookmarkName, 'Член5') then begin
              case komisiya of
              1: Range.Text:= GEK1.Lines[4];
              2: Range.Text:= GEK2.Lines[4];
              3: Range.Text:= GEK3.Lines[4];
              4: Range.Text:= GEK4.Lines[4];
              5: Range.Text:= GEK5.Lines[4];
              6: Range.Text:= GEK6.Lines[4];
              7: Range.Text:= GEK7.Lines[4];
              8: Range.Text:= GEK8.Lines[4];
              end;
            end;

            If CompareBm(BookmarkName, 'Член6') then begin
              case komisiya of
              1: Range.Text:= GEK1.Lines[5];
              2: Range.Text:= GEK2.Lines[5];
              3: Range.Text:= GEK3.Lines[5];
              4: Range.Text:= GEK4.Lines[5];
              5: Range.Text:= GEK5.Lines[5];
              6: Range.Text:= GEK6.Lines[5];
              7: Range.Text:= GEK7.Lines[5];
              8: Range.Text:= GEK8.Lines[5];
              end;
            end;

            if CompareBm(BookmarkName, 'Дата1') then
              Range.Text:= E_Day1.Text;

            if CompareBm(BookmarkName, 'Дата2') then
              Range.Text:= E_Day2.Text;

            if CompareBm(BookmarkName, 'Дата3') then
              Range.Text:= E_Day3.Text;

            if CompareBm(BookmarkName, 'Месяц3') then
              Range.Text:= E_Month3.Text;

            if CompareBm(BookmarkName, 'ЭМП') then
              Range.Text:= EMP.Text;

            if CompareBm(BookmarkName, 'Стац') then
              Range.Text:= Stacionar.Text;

            if CompareBm(BookmarkName, 'Пол') then
              Range.Text:= Poliklinika.Text;

            if CompareBm(BookmarkName, 'СЛР') then
              Range.Text:= SLR.Text;

            if CompareBm(BookmarkName, 'Этап1') then
              Range.Text:= Itog1.Text;

            if CompareBm(BookmarkName, 'ТестОтв') then
              Range.Text:= E_TestCor.Text;

            if CompareBm(BookmarkName, 'Этап2') then
              Range.Text:= Itog2.Text;

            if CompareBm(BookmarkName, 'НомерБилета') then
              Range.Text:= E_NumByl.Text;

            If CompareBm(BookmarkName, 'ХарПоБил') then begin
              case character of
              1: Range.Text:= Great.Text;
              2: Range.Text:= Good.Text;
              3: Range.Text:= Passable.Text;
              4: Range.Text:= Substandard.Text;
              end;
            end;

            if CompareBm(BookmarkName, 'ДопВопр') then
              Range.Text:= M_DopVop.Text;

            If CompareBm(BookmarkName, 'ХарДопВоп') then begin
              case dopchar of
              1: Range.Text:= Dop_Great.Text;
              2: Range.Text:= Dop_Good.Text;
              3: Range.Text:= Dop_Passable.Text;
              4: Range.Text:= Dop_Substandard.Text;
              end;
            end;

            if CompareBm(BookmarkName, 'Этап3') then
              Range.Text:= Itog3.Text;

            if CompareBm(BookmarkName, 'Итог') then
              Range.Text:= ItogGIA.Text;

            if CompareBm(BookmarkName, 'Коммент') then
              Range.Text:= M_Comm.Text;

            if CompareBm(BookmarkName, 'Председатель') then
              Range.Text:= E_Predsedatel.Text;

            if CompareBm(BookmarkName, 'Секретарь') then
              Range.Text:= E_Secretar.Text;
          end;

          tmp_str := ExtractFileDir(Application.ExeName) + PathDelim +'Протокола ГИА'+ PathDelim +E_Grup.Text;
          if not DirectoryExists(tmp_str) then
              ForceDirectories(tmp_str);
          if not CheckBox1.Checked then
            s := tmp_str + '\'+CB_StudFIO.Text+'.doc'
          else
            s := tmp_str + '\'+E_StudFIOManual.Text+'.doc';

          WordApp.ActiveDocument.Saveas(s); // сохраняет файл где нужно
          WordApp.ActiveDocument.Close(True);
          WordApp.Quit;
          WordApp := Unassigned;
          Document := Unassigned;
    end;
    E_ProcProt.Text:= 'Закончил!';
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  E_NumProt.Text:= '';
  CB_StudFIO.ItemIndex:= -1;
  E_StudFIOManual.Text:= '';
  CheckBox1.Checked:= False;
  Button1Click(nil);
  Button2Click(nil);
  Button3Click(nil);
  Button4Click(nil);
  E_NumByl.Text:= '';
  M_DopVop.Lines.Clear;
  Button5Click(nil);
  Button6Click(nil);
  Button7Click(nil);
  M_Comm.Lines.Clear;
  E_ProcProt.text:= '';
end;

procedure TForm1.B_TitulClick(Sender: TObject);
var
  TempleateFileName: string;
  WordApp, Document: OLEVariant;
  NameDOT, tmp_str, s: string;
  BookmarkName: string;
  Range, varcol: OLEVariant;
  i,j: integer;

  function CompareBm(ABmName: string; const AName: string): boolean;
  // проверяет наличие закладок
  var
    i: integer;
  begin
    i := Pos('__', ABmName);
    If i > 0 then
      Delete(ABmName, i, Length(ABmName) - i + 1);
    Result := SameText(ABmName, AName);
  end;

begin
    NameDOT := '\TitulGIA.dot';
    E_ProcTitul.Text:= '';
      for j := 0 to M_FIOTitul.Lines.Count - 1 do begin
        E_ProcTitul.Text:= 'Создаю: '+M_FIOTitul.Lines[j];
        If NameDOT <> '' then begin
          TempleateFileName := ExtractFileDir(Application.ExeName) + PathDelim + 'Templates' + NameDOT; // шаблон

          try
            WordApp := CreateOleObject('Word.Application');
          except
            On E: Exception do
            begin
              MessageBox(Application.Handle,
                PChar('Не удалось запустить MS Word!'#13#10 + E.Message), 'Ошибка',
                MB_OK + MB_ICONERROR + MB_SYSTEMMODAL);
              Exit;
            end;
          end;
          Document := WordApp.Documents.add(TempleateFileName);


            for i := Document.Bookmarks.Count downto 1 do
              begin // начинает перебирать все закладки в шаблоне
                BookmarkName := Document.Bookmarks.Item(i).Name;
                Range := Document.Bookmarks.Item(i).Range;
                If CompareBm(BookmarkName, 'Группа') then
                  Range.Text := E_NumGrupTitul.Text;

                If CompareBm(BookmarkName, 'Номер') then
                  Range.Text := M_NumTel.Lines[j];

                If CompareBm(BookmarkName, 'ФИО') then
                  Range.Text := M_FIOTitul.Lines[j];
              end;
              tmp_str := ExtractFileDir(Application.ExeName) + PathDelim +'Титульные листы'+ PathDelim +E_NumGrupTitul.Text;
            if not DirectoryExists(tmp_str) then
              ForceDirectories(tmp_str);
            s := tmp_str + '\'+M_FIOTitul.Lines[j]+'.docx';
            WordApp.ActiveDocument.Saveas(s); // сохраняет файл где нужно
            WordApp.ActiveDocument.Close(True);
            WordApp.Quit;
            WordApp := Unassigned;
            Document := Unassigned;
        end;
      end;
      E_ProcTitul.Text:= 'Закончил!';
      M_FIOTitul.Lines.Clear;
      M_NumTel.Lines.Clear;
      E_NumGrupTitul.Text:= IntToStr(StrToInt(E_NumGrupTitul.Text) + 1);
end;

procedure TForm1.B_VypiskaClick(Sender: TObject);
var
  TempleateFileName: string;
  WordApp, Document: OLEVariant;
  NameDOT, tmp_str, s: string;
  BookmarkName: string;
  Range, varcol: OLEVariant;
  i,j: integer;

  function CompareBm(ABmName: string; const AName: string): boolean;
  // проверяет наличие закладок
  var
    i: integer;
  begin
    i := Pos('__', ABmName);
    If i > 0 then
      Delete(ABmName, i, Length(ABmName) - i + 1);
    Result := SameText(ABmName, AName);
  end;

begin
    NameDOT := '\Vypiska.dot';
    E_ProcVyp.Text:= '';
      for j := 0 to M_FIOVyp.Lines.Count - 1 do begin
        E_ProcVyp.Text:= 'Создаю: '+M_FIOVyp.Lines[j];
        If NameDOT <> '' then begin
          TempleateFileName := ExtractFileDir(Application.ExeName) + PathDelim + 'Templates' + NameDOT; // шаблон

          try
            WordApp := CreateOleObject('Word.Application');
          except
            On E: Exception do
            begin
              MessageBox(Application.Handle,
                PChar('Не удалось запустить MS Word!'#13#10 + E.Message), 'Ошибка',
                MB_OK + MB_ICONERROR + MB_SYSTEMMODAL);
              Exit;
            end;
          end;
          Document := WordApp.Documents.add(TempleateFileName);


            for i := Document.Bookmarks.Count downto 1 do
              begin // начинает перебирать все закладки в шаблоне
                BookmarkName := Document.Bookmarks.Item(i).Name;
                Range := Document.Bookmarks.Item(i).Range;
                If CompareBm(BookmarkName, 'ФИО') then
                  Range.Text := M_FIOVyp.Lines[j];

                If CompareBm(BookmarkName, 'диплом') then
                  Range.Text := M_Diplom.Lines[j];
              end;
            if M_Diplom.Lines[j] = 'с отличием' then
              tmp_str := ExtractFileDir(Application.ExeName) + PathDelim +'Выписки ГАК'+ PathDelim +'Краснодипломники'
            else
              tmp_str := ExtractFileDir(Application.ExeName) + PathDelim +'Выписки ГАК'+ PathDelim +E_NumGrup.Text;
            if not DirectoryExists(tmp_str) then
              ForceDirectories(tmp_str);
            if M_Diplom.Lines[j] = 'с отличием' then
              s := tmp_str + '\'+E_NumGrup.Text+'_'+M_FIOVyp.Lines[j]+'.docx'
            else
              s := tmp_str + '\'+M_FIOVyp.Lines[j]+'.docx';
            WordApp.ActiveDocument.Saveas(s); // сохраняет файл где нужно
            WordApp.ActiveDocument.Close(True);
            WordApp.Quit;
            WordApp := Unassigned;
            Document := Unassigned;
        end;
      end;
      E_ProcVyp.Text:= 'Закончил!';
      M_FIOVyp.Lines.Clear;
      M_Diplom.Lines.Clear;
      E_NumGrup.Text:= IntToStr(StrToInt(E_NumGrup.Text) + 1);
end;

procedure TForm1.CB_AnsGoodClick(Sender: TObject);
begin
  if CB_AnsGood.Checked then begin
    CB_AnsOtl.Checked:= False;
    CB_AnsUd.Checked:= False;
    CB_Ansneud.Checked:= False;
    character:= 2;
  end
  else
    character:= 0;
end;

procedure TForm1.CB_AnsneudClick(Sender: TObject);
begin
  if CB_Ansneud.Checked then begin
    CB_AnsOtl.Checked:= False;
    CB_AnsGood.Checked:= False;
    CB_AnsUd.Checked:= False;
    character:= 4;
  end
  else
    character:= 0;
end;

procedure TForm1.CB_AnsOtlClick(Sender: TObject);
begin
  if CB_AnsOtl.Checked then begin
    CB_AnsGood.Checked:= False;
    CB_AnsUd.Checked:= False;
    CB_Ansneud.Checked:= False;
    character:= 1;
  end
  else
    character:= 0;
end;

procedure TForm1.CB_AnsUdClick(Sender: TObject);
begin
  if CB_AnsUd.Checked then begin
    CB_AnsOtl.Checked:= False;
    CB_AnsGood.Checked:= False;
    CB_Ansneud.Checked:= False;
    character:= 3;
  end
  else
    character:= 0;
end;

procedure TForm1.CB_DopAnsGoodClick(Sender: TObject);
begin
  if CB_DopAnsGood.Checked then begin
    CB_DopAnsOtl.Checked:= False;
    CB_DopAnsUd.Checked:= False;
    CB_DopAnsNeud.Checked:= False;
    dopchar:= 2;
  end
  else
    dopchar:= 0;
end;

procedure TForm1.CB_DopAnsNeudClick(Sender: TObject);
begin
  if CB_DopAnsNeud.Checked then begin
    CB_DopAnsOtl.Checked:= False;
    CB_DopAnsGood.Checked:= False;
    CB_DopAnsUd.Checked:= False;
    dopchar:= 4;
  end
  else
    dopchar:= 0;
end;

procedure TForm1.CB_DopAnsOtlClick(Sender: TObject);
begin
  if CB_DopAnsOtl.Checked then begin
    CB_DopAnsGood.Checked:= False;
    CB_DopAnsUd.Checked:= False;
    CB_DopAnsNeud.Checked:= False;
    dopchar:= 1;
  end
  else
    dopchar:= 0;
end;

procedure TForm1.CB_DopAnsUdClick(Sender: TObject);
begin
  if CB_DopAnsUd.Checked then begin
    CB_DopAnsOtl.Checked:= False;
    CB_DopAnsGood.Checked:= False;
    CB_DopAnsNeud.Checked:= False;
    dopchar:= 3;
  end
  else
    dopchar:= 0;
end;

procedure TForm1.CB_GEK1Click(Sender: TObject);
begin
  if CB_GEK1.Checked then begin
    CB_GEK2.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK7.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 1;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK2Click(Sender: TObject);
begin
  if CB_GEK2.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK7.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 2;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK3Click(Sender: TObject);
begin
  if CB_GEK3.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK2.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK7.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 3;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK4Click(Sender: TObject);
begin
  if CB_GEK4.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK2.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK7.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 4;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK5Click(Sender: TObject);
begin
  if CB_GEK5.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK2.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK7.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 5;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK6Click(Sender: TObject);
begin
  if CB_GEK6.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK2.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK7.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 6;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK7Click(Sender: TObject);
begin
  if CB_GEK7.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK2.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK8.Checked:= False;
    komisiya:= 7;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CB_GEK8Click(Sender: TObject);
begin
  if CB_GEK8.Checked then begin
    CB_GEK1.Checked:= False;
    CB_GEK2.Checked:= False;
    CB_GEK3.Checked:= False;
    CB_GEK4.Checked:= False;
    CB_GEK5.Checked:= False;
    CB_GEK6.Checked:= False;
    CB_GEK7.Checked:= False;
    komisiya:= 8;
  end
  else
    komisiya:= 0;
end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked then
    E_StudFIOManual.Enabled:= True
  else
    E_StudFIOManual.Enabled:= False;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var
  Ini: Tinifile;
begin
  Ini:=TiniFile.Create(GetCurrentDir+PREFORM+PathDelim+'const.ini');
  with Ini do begin
    WriteString('Spec', 'Spec', E_Spec.Text);
    WriteString('PredsedatelSecretar', 'PredsedatelFIO', E_PredsedatelFIO.Text);
    WriteInteger('DateDay', 'Day1', StrToInt(E_Day1.Text));
    WriteInteger('DateDay', 'Day2', StrToInt(E_Day2.Text));
    WriteString('СharacterAnswer', 'Great', Great.Lines.Text);
    WriteString('СharacterAnswer', 'Good', Good.Lines.Text);
    WriteString('СharacterAnswer', 'Passable', Passable.Lines.Text);
    WriteString('СharacterAnswer', 'Substandard', Substandard.Lines.Text);
    WriteString('СharacterDopAnswer', 'Great', Dop_Great.Lines.Text);
    WriteString('СharacterDopAnswer', 'Good', Dop_Good.Lines.Text);
    WriteString('СharacterDopAnswer', 'Passable', Dop_Passable.Lines.Text);
    WriteString('СharacterDopAnswer', 'Substandard', Dop_Substandard.Lines.Text);
    WriteString('PredsedatelSecretar', 'PredsedatelSocr', E_Predsedatel.Text);
    WriteString('PredsedatelSecretar', 'SecretarSocr', E_Secretar.Text);
    WriteString('DateDay', 'Month', E_Month3.Text);
    Free;
  end;

  GEK1.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK1.txt');
  GEK2.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK2.txt');
  GEK3.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK3.txt');
  GEK4.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK4.txt');
  GEK5.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK5.txt');
  GEK6.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK6.txt');
  GEK7.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK7.txt');
  GEK8.Lines.SaveToFile(GetCurrentDir+PREFORM+PathDelim+'GEK8.txt');
end;

procedure TForm1.FormShow(Sender: TObject);
var
  Ini: Tinifile;
begin
  CB_StudFIO.Items.Clear;
  CB_StudFIO.Items.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'StudRodP.txt');

  Ini:=TiniFile.Create(GetCurrentDir+PREFORM+PathDelim+'const.ini');
  E_Spec.Text:= Ini.ReadString('Spec', 'Spec', '');
  E_PredsedatelFIO.Text:= Ini.ReadString('PredsedatelSecretar', 'PredsedatelFIO', '');
  E_Day1.Text:= Ini.ReadString('DateDay', 'Day1', '');
  E_Day2.Text:= Ini.ReadString('DateDay', 'Day2', '');
  Great.Lines.Add(Ini.ReadString('СharacterAnswer', 'Great', ''));
  Good.Lines.Add(Ini.ReadString('СharacterAnswer', 'Good', ''));
  Passable.Lines.Add(Ini.ReadString('СharacterAnswer', 'Passable', ''));
  Substandard.Lines.Add(Ini.ReadString('СharacterAnswer', 'Substandard', ''));
  Dop_Great.Lines.Add(Ini.ReadString('СharacterDopAnswer', 'Great', ''));
  Dop_Good.Lines.Add(Ini.ReadString('СharacterDopAnswer', 'Good', ''));
  Dop_Passable.Lines.Add(Ini.ReadString('СharacterDopAnswer', 'Passable', ''));
  Dop_Substandard.Lines.Add(Ini.ReadString('СharacterDopAnswer', 'Substandard', ''));
  E_Predsedatel.Text:= Ini.ReadString('PredsedatelSecretar', 'PredsedatelSocr', '');
  E_Secretar.Text:= Ini.ReadString('PredsedatelSecretar', 'SecretarSocr', '');
  E_Month3.Text:= Ini.ReadString('DateDay', 'Month', '');
  Ini.Free;

  GEK1.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK1.txt');
  GEK2.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK2.txt');
  GEK3.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK3.txt');
  GEK4.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK4.txt');
  GEK5.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK5.txt');
  GEK6.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK6.txt');
  GEK7.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK7.txt');
  GEK8.Lines.LoadFromFile(GetCurrentDir+PREFORM+PathDelim+'GEK8.txt');
end;

procedure TForm1.ScrollBox1MouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  (Sender as TScrollBox).VertScrollBar.Position:= (Sender as TScrollBox).VertScrollBar.Position + 35;
end;

procedure TForm1.ScrollBox1MouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  (Sender as TScrollBox).VertScrollBar.Position:= (Sender as TScrollBox).VertScrollBar.Position - 35;
end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
  ScrollBox1.VertScrollBar.Position:= 0;
end;

end.

