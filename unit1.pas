unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, XMLPropStorage, ComCtrls, variants, comobj, ShellApi, windows;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    OpenDialog1: TOpenDialog;
    ProgressBar1: TProgressBar;
    SelectDirectoryDialog1: TSelectDirectoryDialog;
    XMLPropStorage1: TXMLPropStorage;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

const
  dir = 'I:\';
  Days = 6;
  Les = 7;

var
  Form1: TForm1;
  Excel: OleVariant;


implementation

{$R *.lfm}
procedure CreateTemplate(Sheet: OleVariant);
var
  i, j: integer;
begin
  j := 2;
  for i:=1 to Les do
    Sheet.Cells[i,j] := i;

  j:= 1;
  for i := 1 to  Days do
    begin
      Sheet.Range[Sheet.Cells[(i-1)*Les+1,j], Sheet.Cells[i*Les,j]].MergeCells:=True;
      if i > 1 then
        Sheet.Range[Sheet.Cells.Item[1,j+1], Sheet.Cells.Item[Les,j+1]].Copy(
          Sheet.Range[Sheet.Cells[(i-1)*Les+1,j+1], Sheet.Cells[i*Les,j+1]]);
    end;
end;

procedure Rasp(ResSheet, Sheet: OleVariant; kab :string; StartRow, StartCol, FirstLesson: integer);
var
  i, j, z, Day, LesNum: integer;
begin
  ResSheet.Columns[3].ColumnWidth := 25;
  ResSheet.Columns[4].ColumnWidth := 25;
  ResSheet.Columns[5].ColumnWidth := 25;

  Day := -1;
  z:=3;
  i:=StartRow;
  j:=StartCol;

  while (i<51) do
  begin
    if (Pos(' пара', Sheet.cells[i,j].text)>0) then
      begin
        LesNum:=StrToInt(Copy(Sheet.cells[i,j].text, 1, 1));
        //ShowMessage(IntToStr(LesNum));
        if (LesNum = FirstLesson) then
          begin
            inc(Day);
            ResSheet.Cells[Day*Les + LesNum,1]:= Sheet.cells[i,1].text ;
            ResSheet.Cells[Day*Les + LesNum,1].Orientation := 90;
            ResSheet.Cells[Day*Les + LesNum,1].HorizontalAlignment := -4108;
            ResSheet.Cells[Day*Les + LesNum,1].VerticalAlignment := -4108;
          end;
        inc(j);
        while (j<100) do
          begin
            if (Sheet.cells[i,j].Text <>'') and
               ((Pos(kab, Sheet.cells[i,j].Text) > 0)) then
              begin
                while (ResSheet.Cells[Day*Les + LesNum, z].Text <> '') do
                  inc(z);
                ResSheet.Cells[Day*Les + LesNum, z]:=Sheet.cells[i,j].Text;
                inc(z);
              end;
            inc(j);
          end;
        z:=3;
        inc(i);
        j:=StartCol;

      end
    else
    begin
      z:=3;
      inc(i);
      j:=StartCol;
    end;
  end;
end;

{ TForm1 }



procedure TForm1.Button2Click(Sender: TObject);
var
  Excel, SheetBak, SheetMag, ResSheet: OleVariant;
  FileName: widestring;
begin


  if (LabeledEdit2.Text = '') or  (LabeledEdit3.Text = '')
    then
      begin
        ShowMessage('Укажите оба файла');
        exit;
      end
  else
  begin
    if not(FileExists(LabeledEdit2.Text)) then
    begin
      ShowMessage('Файл ' + ExtractFileName(LabeledEdit2.Text)+ ' не найден');
      LabeledEdit2.Text:='';
      OpenDialog1.FileName:='';
      OpenDialog1.InitialDir:='';
      Exit;
    end;

  if not(FileExists(LabeledEdit3.Text)) then
      begin
        ShowMessage('Файл ' + ExtractFileName(LabeledEdit3.Text)+ ' не найден');
        LabeledEdit3.Text:='';
        OpenDialog1.FileName:='';
        OpenDialog1.InitialDir:='';
        Exit;
      end;
  end;

   if  not(((ExtractFileExt(LabeledEdit2.Text) <> '.xls')
    or (ExtractFileExt(LabeledEdit2.Text) <> '.xlsx')) or
    ((ExtractFileExt(LabeledEdit3.Text) <> '.xls')
    or (ExtractFileExt(LabeledEdit3.Text) <> '.xlsx')))
    then
      begin
        ShowMessage('Файлы не Excel расшинерия');
        Exit;
      end;

   if (LabeledEdit1.Text = '') then
     begin
       ShowMessage('Укажите номер кабинета или фамилию преподавателя');
       Exit;
     end;
   SelectDirectoryDialog1.Title:='Выберите папку для сохранения результата';
   ProgressBar1.Visible:=True;
   Application.ProcessMessages;
     if SelectDirectoryDialog1.Execute then
       begin
         Excel:=CreateOleObject('Excel.Application');
         Excel.Visible:=False;
         //Excel.DisableAlerts:=true;
         Excel.Workbooks.Add;
         Excel.Workbooks.Open(WideString(LabeledEdit2.Text));
         Excel.Workbooks.Open(WideString(LabeledEdit3.Text));
         ResSheet:=Excel.Workbooks[1].WorkSheets[1];
         SheetBak:=Excel.Workbooks[2].WorkSheets[1];
         SheetMag:=Excel.Workbooks[3].WorkSheets[1];
         FileName:=WideString(SelectDirectoryDialog1.FileName +'\Kab'+LabeledEdit1.Text+'.xlsx');

         CreateTemplate(ResSheet);
         Rasp(ResSheet, SheetBak, LabeledEdit1.Text, 8, 2, 1);
         Rasp(ResSheet, SheetMag, LabeledEdit1.Text, 10, 2, 4);
         SelectDirectoryDialog1.InitialDir:=SelectDirectoryDialog1.FileName;

         Excel.WorkBooks[1].SaveAs(FileName);
         ResSheet:=Unassigned;
         SheetBak:=Unassigned;
         SheetMag:=Unassigned;
         Excel.WorkBooks.Close;


         //ShowMessage('Выполнено!');
         Excel.Workbooks.Open(FileName);
         Excel.Visible:=True;

         //Excel.Quit;
         Excel:=Unassigned;
         //ShowMessage(FileName);
         //ShellExecute(handle,'open',PChar(SelectDirectoryDialog1.FileName +'\Kab'+LabeledEdit1.Text+'.xlsx'), '','',SW_MAXIMIZE);
       end;
     ProgressBar1.Visible:=False;
  end;


procedure TForm1.Button3Click(Sender: TObject);
begin
  OpenDialog1.FileName:='';
  if OpenDialog1.Execute then
    if OpenDialog1.FileName <> ''
       then LabeledEdit3.Text := OpenDialog1.FileName;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  OpenDialog1.FileName:='';
  if OpenDialog1.Execute then
    if OpenDialog1.FileName <> ''
       then LabeledEdit2.Text := OpenDialog1.FileName;
end;

end.

