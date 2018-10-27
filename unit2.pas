unit Unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Windows;

type

  { TForm2 }

  TForm2 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    year: TComboBox;
    month: TComboBox;
    day: TComboBox;
    procedure Button1Click(Sender: TObject);
  private

  public

  end;

const
  SENIOR_1 = 0;
  SENIOR_2 = 1;
  SENIOR_3 = 2;

var
  Form2: TForm2;
  yearString,monthString,dayString: String;

implementation

{$R *.lfm}

{ TForm2 }

procedure TForm2.Button1Click(Sender: TObject);
begin
  yearString := PChar(year.Text);
  monthString := PChar(month.Text);
  dayString := PChar(day.Text);
  ShellExecute(0,'open','download_with_date.exe',PChar('--Grade 0 --Pattern 0 --Year '
    +yearString+' --Month '+monthString+' --Day '+dayString),'',SW_SHOW);
  ShellExecute(0,'open','download_with_date.exe',PChar('--Grade 0 --Pattern 1 --Year '
    +yearString+' --Month '+monthString+' --Day '+dayString),'',SW_SHOW);
  form2.Close;
end;

end.

