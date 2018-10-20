unit Unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, CheckLst,
  StdCtrls;

type

  { TForm2 }

  TForm2 = class(TForm)
    Button1: TButton;
    grade: TComboBox;
    procedure Button1Click(Sender: TObject);
  private

  public
    method : integer;
  end;

const
  SENIOR_1 = 0;
  SENIOR_2 = 1;
  SENIOR_3 = 2;

var
  Form2: TForm2;

implementation

{$R *.lfm}

{ TForm2 }

procedure TForm2.Button1Click(Sender: TObject);
begin
  form2.Close;
end;

end.

