unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, variants, comobj;

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
    SelectDirectoryDialog1: TSelectDirectoryDialog;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
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
Function CreateTemplate(Directory, kab: string):string;
var
  Excel, Sheet: OleVariant;
  i, j: integer;
begin
  Excel:=CreateOleObject('Excel.Application');
  Excel.Visible:=False;
  Excel.Workbooks.Add;
  Sheet:=Excel.Workbooks[1].WorkSheets[1];

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

  Excel.WorkBooks[1].SaveAs(WideString(Directory+'Kab'+kab+'.xlsx'));
  Excel.Quit;
  Excel:=Unassigned;
  Sheet:=Unassigned;
  CreateTemplate:=Directory+'Kab'+kab+'.xlsx';
end;

procedure Rasp(RaspFileName, ResFileName, kab :string; StartRow, StartCol, FirstLesson: integer);
var
  Excel, Books, Sheet, ResSheet: OleVariant;
  i, j, k, z, Day, LesNum: integer;
begin
  Excel:=CreateOleObject('Excel.Application');
  Excel.Visible:=False;
  Excel.Workbooks.Open(WideString(ResFileName));
  Excel.Workbooks.Open(WideString(RaspFileName));
  ResSheet:=Excel.WorkBooks[1].Sheets[1];
  Sheet := Excel.WorkBooks[2].Sheets[1];

  ResSheet.Columns[3].ColumnWidth := 25;
  ResSheet.Columns[4].ColumnWidth := 25;
  ResSheet.Columns[5].ColumnWidth := 25;

  Day := -1;
  k:=1;
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
  Excel.WorkBooks[1].SaveAs(WideString(ResFileName));
  Excel.Quit;
  Excel:=Unassigned;
  Sheet:=Unassigned;
  ResSheet:=Unassigned;
end;

{ TForm1 }



procedure TForm1.Button2Click(Sender: TObject);
var
  FileName: string;
begin
  if (LabeledEdit1.Text <> '') and ((ExtractFileExt(LabeledEdit2.Text) = '.xls')
    or (ExtractFileExt(LabeledEdit2.Text) = '.xlsx')) and
    ((ExtractFileExt(LabeledEdit3.Text) = '.xls')
    or (ExtractFileExt(LabeledEdit3.Text) = '.xlsx')) then
      begin
        SelectDirectoryDialog1.Title:='Выберите папку для сохранения результата';
        if SelectDirectoryDialog1.Execute then
          begin
            FileName := CreateTemplate(SelectDirectoryDialog1.FileName, LabeledEdit1.Text);
            Rasp(LabeledEdit2.Text, FileName, LabeledEdit1.Text, 8, 2, 1);
            Rasp(LabeledEdit3.Text, FileName, LabeledEdit1.Text, 10, 2, 4);
          end;
      end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  OpenDialog1.FileName:='';
  if OpenDialog1.Execute then
    if OpenDialog1.FileName <> ''
       then LabeledEdit3.Text := OpenDialog1.FileName;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  OpenDialog1.FileName:='';
  if OpenDialog1.Execute then
    if OpenDialog1.FileName <> ''
       then LabeledEdit2.Text := OpenDialog1.FileName;
end;

end.

