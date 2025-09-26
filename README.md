# Phone-Book-In-Python
A simple phone book in python, terminal based that handles excel data loading and writing

In case you want to use it for your Business ERP, use the IDE of your ERP (using Pascal/Basic). For example,
Unit 1:
uses
  Classes, Graphics, Controls, Forms, Dialogs, Unit2;

var
  MainForm: TForm2;
begin
  MainForm := TForm2.Create(Application);
  MainForm.Show;
end;

Unit 2:
{$FORM TForm2, Unit2.sfm}

uses
  Classes, Graphics, Controls, Forms, Dialogs;

procedure Form2Create(Sender: TObject);
begin
   ExecuteProcess('C:\Be.POSI\book\book.exe', TRUE);
end;

begin
end;

The application runs searching for your Hosted File that the ERP uses and then for the exe.
