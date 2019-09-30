# PreviewHandler-VCL-
Simple Preview Handler Component for VCL applications on Win32 by Uğur PARLAYAN

###### for example

```pascal
unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ShellCtrls, ExtCtrls, StdCtrls, dxPreviewHandler, Vcl.FileCtrl, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore, dxSkinVS2010, cxGroupBox;

type
  TForm1 = class(TForm)
    Panel1            : TPanel;
    Panel2            : TPanel;
    DirectoryListBox1 : TDirectoryListBox;
    FileListBox1      : TFileListBox;
    DriveComboBox1    : TDriveComboBox;
    Splitter1         : TSplitter;
    Splitter2         : TSplitter;
    dxPreviewHandler1 : TdxPreviewHandler;
    procedure FormCreate(Sender: TObject);
    procedure FileListBox1Click(Sender: TObject);
    procedure Panel1Resize(Sender: TObject);
  private
  public
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  DriveComboBox1.DirList := DirectoryListBox1;
  DirectoryListBox1.FileList := FileListBox1;
  dxPreviewHandler1.Control := Panel1;
end;

procedure TForm1.FileListBox1Click(Sender: TObject);
begin
  if (FileExists(FileListBox1.FileName, True) = TRUE) then begin
      dxPreviewHandler1.FileName := FileListBox1.FileName;
      dxPreviewHandler1.Preview;
  end;
end;

procedure TForm1.Panel1Resize(Sender: TObject);
begin
  dxPreviewHandler1.Resize;
end;

end.
```
