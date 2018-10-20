unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Unit2;

type

  { TForm1 }

  TForm1 = class(TForm)
    DeleteTimerTask: TButton;
    GetDailyExcel: TButton;
    CreateTimerTask: TButton;
    GetWeeklyExcel: TButton;
    procedure Button1Click(Sender: TObject);
    procedure CreateTimerTaskClick(Sender: TObject);
    procedure DeleteTimerTaskClick(Sender: TObject);
    procedure GetWeeklyExcelClick(Sender: TObject);
    procedure GetDailyExcelClick(Sender: TObject);
  private

  public

  end;

const
  CREATE_TASK = 0;
  GET_DAILY_EXCEL = 1;
  GET_WEEKLY_EXCEL = 2;
var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin

end;

procedure TForm1.CreateTimerTaskClick(Sender: TObject);
begin
  form2.method:=CREATE_TASK;
  form2.Show;
end;

procedure TForm1.DeleteTimerTaskClick(Sender: TObject);
begin

end;

procedure TForm1.GetWeeklyExcelClick(Sender: TObject);
begin
  form2.method:=GET_WEEKLY_EXCEL;
  form2.Show;
end;

procedure TForm1.GetDailyExcelClick(Sender: TObject);
begin
  form2.method:=GET_DAILY_EXCEL;
  form2.Show;
end;

end.

