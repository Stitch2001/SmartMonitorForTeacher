unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Unit2, Windows, Unit3;

type

  { TForm1 }

  TForm1 = class(TForm)
    getExcelByDate: TButton;
    getDailyExcel: TButton;
    createTimerTask: TButton;
    About: TButton;
    getWeeklyExcel: TButton;
    procedure AboutClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure createTimerTaskClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure getExcelByDateClick(Sender: TObject);
    procedure getWeeklyExcelClick(Sender: TObject);
    procedure getDailyExcelClick(Sender: TObject);
  private

  public

  end;


var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin

end;

procedure TForm1.AboutClick(Sender: TObject);
begin
  form3.Show;
end;

procedure TForm1.createTimerTaskClick(Sender: TObject);
begin
  ShowMessage('Win7以上系统请以管理员模式运行本软件，否则此操作无效');
  ShellExecute(0,'open','cmd.exe',
    '/c reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v AUTORUN /t ' +
    'REG_SZ /d "daily_download.exe --Grade 0 --Pattern 0" /f','',SW_SHOW);
  ShellExecute(0,'open','cmd.exe',
    '/c reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v AUTORUN /t ' +
    'REG_SZ /d "daily_download.exe --Grade 0 --Pattern 1" /f','',SW_SHOW);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  ShellExecute(0,'open','check_updates1.exe','','',SW_SHOW);
end;

procedure TForm1.getExcelByDateClick(Sender: TObject);
begin
  form2.Show;
end;

procedure TForm1.getWeeklyExcelClick(Sender: TObject);
begin
  ShellExecute(0,'open','weekly_download.exe','--Grade 0','',SW_SHOW);
end;

procedure TForm1.getDailyExcelClick(Sender: TObject);
begin
  ShellExecute(0,'open','daily_download.exe','--Grade 0 --Pattern 0','',SW_SHOW);
  ShellExecute(0,'open','daily_download.exe','--Grade 0 --Pattern 1','',SW_SHOW);
end;

end.

