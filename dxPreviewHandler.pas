/// <summary>
///  Windows işletim sisteminin IPreviewHandler apisini kullanarak döküman önizleme yapılabilmesini sağlayan
///  bileşendir.
/// </summary>
///  <example>
///    Örnek kullanım için aşağıdaki kodu inceleyebilirsiniz
///  <code lang="Delphi">
///  unit Unit1;
///
///  interface
///
///  uses
///    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
///    Dialogs, ComCtrls, ShellCtrls, ExtCtrls, StdCtrls, dxPreviewHandler, Vcl.FileCtrl, cxGraphics, cxControls,
///    cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore, dxSkinVS2010, cxGroupBox;
///
///  type
///    TForm1 = class(TForm)
///      Panel1            : TPanel;
///      Panel2            : TPanel;
///      DirectoryListBox1 : TDirectoryListBox;
///      FileListBox1      : TFileListBox;
///      DriveComboBox1    : TDriveComboBox;
///      Splitter1         : TSplitter;
///      Splitter2         : TSplitter;
///      dxPreviewHandler1 : TdxPreviewHandler;
///      procedure FormCreate(Sender: TObject);
///      procedure FileListBox1Click(Sender: TObject);
///      procedure Panel1Resize(Sender: TObject);
///    private
///    public
///    end;
///
///  var
///    Form1: TForm1;
///
///  implementation
///
///  {$R *.dfm}
///
///  procedure TForm1.FormCreate(Sender: TObject);
///  begin
///    DriveComboBox1.DirList := DirectoryListBox1;
///    DirectoryListBox1.FileList := FileListBox1;
///    dxPreviewHandler1.Control := Panel1;
///  end;
///
///  procedure TForm1.FileListBox1Click(Sender: TObject);
///  begin
///    if (FileExists(FileListBox1.FileName, True) = TRUE) then begin
///        dxPreviewHandler1.FileName := FileListBox1.FileName;
///        dxPreviewHandler1.Preview;
///    end;
///  end;
///
///  procedure TForm1.Panel1Resize(Sender: TObject);
///  begin
///    dxPreviewHandler1.Resize;
///  end;
///
///  end.
///  </code>
///  </example>
/// <remarks>
///  Copyright © Uğur PARLAYAN (01.05.2013) Lütfen kaynak belirterek kullanınız.
/// </remarks>
unit dxPreviewHandler;
 
interface
 
uses
 StdCtrls,
 Registry, ComObj, ActiveX, AxCtrls, ShlObj,
 Windows, Classes, Controls, SysUtils, Dialogs;
 
type
 /// <summary>
 ///  Detaylarına şu linkten ulaşabilirsiniz;
 ///  <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb761818(v=vs.85).aspx" />
 /// </summary>
 IInitializeWithFile   = interface(IUnknown) ['{b7d14566-0509-4cce-a71f-0a554233bd9b}']
   function Initialize(pszFilePath: LPWSTR; grfMode: DWORD):HRESULT; stdcall;
 end;
 
 /// <summary>
 ///  Detaylarına şu linkten ulaşabilirsiniz;
 ///  <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb761810(v=vs.85).aspx" />
 /// </summary>
 IInitializeWithStream = interface(IUnknown) ['{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}']
   function Initialize(pstream: IStream; grfMode: DWORD): HRESULT; stdcall;
 end;
 
 /// <summary>
 ///  Detaylarına şu linkten ulaşabilirsiniz;
 ///  <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/cc144143(v=vs.85).aspx" />
 /// </summary>
 IPreviewHandler       = interface(IUnknown) ['{8895b1c6-b41f-4c1c-a562-0d564250836f}']
   function SetWindow(hwnd: HWND; var RectangleRef: TRect): HRESULT; stdcall;
   function SetRect(var RectangleRef: TRect): HRESULT; stdcall;
   function DoPreview(): HRESULT; stdcall;
   function Unload(): HRESULT; stdcall;
   function SetFocus(): HRESULT; stdcall;
   function QueryFocus(phwnd: HWND): HRESULT; stdcall;
   function TranslateAccelerator(PointerToWindowMessage: MSG): HRESULT; stdcall;
 end;

 TdxPreviewHandler = class(TComponent)
 private
   FFileName : string;           // Gösterilecek dosyanın yolu ve adı bu değişkende saklanacaktır.
   FHandler  : IPreviewHandler;  // Bu arabirim aracılığıyla api çağırılacaktır.
   FControl  : TWinControl;      // Bu nesne üzerinde önizleme yapılacaktır.
   FClassID  : String;           // Bir TGUID değeri barındırmaktadır.
   FKip      : String;           // PreviewFile prosedüründe kullanılan en son kipi tutar.

   /// <summary>
   ///  PreviewFile prosedürünü tetikleyecek şekilde görüntülenmek istenen dosyanın tam yolunu ve adını vermemizi sağlar.
   /// </summary>
   procedure SetFileName(const Value: string);
 
   /// <summary>
   ///  FControl üzerine binecek şekilde belgeyi görüntüler.
   /// </summary>
   procedure PreviewFile;
 
   /// <summary>
   ///  Dosyanın uzantısını okur.
   /// </summary>
   function GetFileExt: String;
 
   /// <summary>
   ///  Dosya uzantısının Registry'deki Preview Handler'inin sınıf kimliğini okur.
   /// </summary>
   function GetClassID: String;
 protected
 public
   /// <summary>
   ///  Klasik Create yordamından farklı olarak nesnenin runtime anında bir preview alanına doğrudan bağlanarak
   ///  üretilmesini sağlar.
   /// </summary>
   /// <param name="aOwner">
   ///  TComponent türünden Parent nesne kastedilmektedir.
   /// </param>
   /// <param name="aControl">
   ///  TWinControl türünden bir nesne, mesela TPanel veya (DevEx'in) TcxGroupBox gibi nesneleri kastedilmektedir.
   /// </param>
   constructor Create(aOwner: TComponent; aControl: TWinControl); overload;
 
   /// <summary>
   ///  Nesne yok edilirken kullandığı IPreviewHandler bağlantısını da yok etmelidir. IPreviewHandler.Unload
   ///  metoduyla birlikte nesne yok edilir.
   /// </summary>
   destructor Destroy; override;
 
   /// <summary>
   ///  IPreviewHandler.Unload medorunu çağırır.
   /// </summary>
   procedure Bosalt;
 
   /// <summary>
   ///  FControl nesnesinin üzerine binecek şekilde belgeyi ekranda görüntülemek için PreviewFile prosedürünü çağırır.
   /// </summary>
   procedure Preview;
 
   /// <summary>
   ///  Bileşen dışından da referans azaltımını kullanabilmek için IPreviewHandler._Release metodunu çağırır.
   /// </summary>
   procedure _DelRef; deprecated; { Önerilmez... }
 
   /// <summary>
   ///  Dökümanın gösterileceği bölgenin yeniden hesaplanmasında kullanılır.
   /// </summary>
   /// <remarks>
   ///  FControl nesnesinin resize olayı tetiklenirken bu prosedür çağırılmalıdır.
   /// </remarks>
   procedure Resize;
 
 published
   /// <summary>
   ///  Dosyanın uzantısını verir.
   /// </summary>
   property FileExt  : String      read GetFileExt;
 
   /// <summary>
   ///  Dosyanın path'iyle birlikte tam adını verir. Örn: "C:\depo\test.pdf" gibi...
   /// </summary>
   property FileName : String      read FFileName      write SetFileName;
 
   /// <summary>
   ///  Registry'de dosya uzantısını açacak olan Preview Handler nesnesinin sınıf kimliğini verir.
   /// </summary>
   property ClassID  : String      read FClassID; // Registry'den okuma işlemini GetClassID fonksiyonu gerçekleştirir. Burada değişken kullanılmasının sebebi Registry trafiğini azaltmak içindir.
 
   /// <summary>
   ///  Belge bu nesnenin üzerinde görüntülenir.
   /// </summary>
   property Control  : TWinControl read FControl       write FControl;
 end;

const
    E_PREVIEWHANDLER_DRM_FAIL   = $86420001;
    E_PREVIEWHANDLER_NOAUTH     = $86420002;
    E_PREVIEWHANDLER_NOTFOUND   = $86420003;
    E_PREVIEWHANDLER_CORRUPT    = $86420004;
 
procedure Register;

implementation

procedure Register;
begin
 RegisterComponents('UpDevEx', [TdxPreviewHandler]);
end;

{ TdxPreviewHandler }
 
constructor TdxPreviewHandler.Create(aOwner: TComponent; aControl: TWinControl);
begin
 inherited Create(aOwner);
 Control := aControl;
end;
 
destructor TdxPreviewHandler.Destroy;
begin
 { FHandler interface tarafından halâ kullanıldığı için yok etme işlemi interface tarafından gerçekleştiriliyor... Ayrıca bir destroy, free ve nil sekansına girilmesine gerek yok... }
 inherited Destroy;
end;
 
procedure TdxPreviewHandler.Bosalt;
begin
 if (Assigned(FHandler) = TRUE) then begin
     FHandler.Unload;
 end;
end;
 
procedure TdxPreviewHandler._DelRef;
begin
 if (Assigned(FHandler) = TRUE) then begin
     FHandler._Release;
 end;
end;
 
function TdxPreviewHandler.GetClassID: String;
var
 aRegistry  :  TRegistry;
 aExtens    ,
 aRegPath   :  String;
begin
 Result := '';
 aRegistry  := TRegistry.Create();
 try
   aRegistry.RootKey := HKEY_CLASSES_ROOT;
   aExtens  := GetFileExt; { mesela '.pdf' veya '.docx' gibi... }
   aRegPath := aExtens + '\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}';
   if (aRegistry.KeyExists(aRegPath) = TRUE) then begin
       aRegistry.Access := KEY_READ; { 32 ve 64 bit windows işletim sistemlerinde registry'e erişim farklı kaynaklardan gerçekleşiyor... }
       aRegistry.OpenKey(aRegPath, FALSE);
       Result := aRegistry.ReadString(''); { Varsayılan değeri okur. Bu bileşen için bu kadarı yeterlidir... }
       aRegistry.CloseKey;
   end;
 finally
   FreeAndNil(aRegistry);
 end;
 FClassID := Result;
end;
 
function TdxPreviewHandler.GetFileExt: String;
begin
 Result := ExtractFileExt(FFileName);
end;
 
procedure TdxPreviewHandler.PreviewFile;
var
 aGUID             : TGUID;
 aRect             : TRect;
 FileInit          : IInitializeWithFile;
 StreamInit        : IInitializeWithStream;
 FS                : TFileStream;
 SA                : IStream;
 TamponBellek      : TMemoryStream;
 Mesaj             : String;
 X: Integer;
begin
 FKip := ''; // Başlangıç değeridir.
 
 { Ön izlemenin yapılacağı konteyner bir nesne yoksa doğrudan çıkılır... }
 if (Assigned(FControl) = False) then Exit;
 
 FClassID := GetClassID;
 
 { Dosya uzantısını açacak bir preview handler ID'si yok demektir... }
 if (FClassID = '') then Exit;
 
 {
   FHandler oluşturulmuşsa şu an bellekte bir dosya preview olarak gösteriliyor demektir. Bu durumda API 2 yoldan birini kendisi seçecektir;
   1) Ya yeni bir FHandler arabirimi üretip yeni belgeyi ona yükleyecek,
   2) Ya da mevcut FHandler arabirimindeki belgeyi bellekten kaldırıp yenisini yükleyecek...
   IPreviewHandler.Unload metodu bize 2. seçeneği uygulama şansı veriyor. Böylece arka taraftaki Referans sayımlarında
   mevcut referansı azaltmakla, yani IPreviewHandler._Release ile uşraşmak zorunda kalmıyoruz. Zira _Release'nin
   RefCount değeri 0'a ulaştığında arabirim kendisini destroy ediyor dolayısıyla arabirimi yeniden üretmek
   gerekebiliyor, çok karışık o nedenle buraya hiç girmemek en iyi çözüm bence...
   Yani özetle unload metodu bellekteki IPreview arabirimini yok etmeden sadece dosya bağlantısını kesip ilgili
   dosyanın bellekteki yerini sıfırlıyor ve yeni bir dosya yükleyebilmemize olanak tanıyor.
   Bunun için BOSALT adlı bir prosedür ekledim.
 }
 if Assigned(FHandler) then FHandler.Unload;
 
 aGUID    := StringToGUID(FClassID);
 FHandler := CreateComObject(aGUID) as IPreviewHandler;
 
 { İşletim sisteminde Dosya uzantısını açabilecek tanımlı bir Preview Handler "nesnesi" yok demektir... Dolayısıyla çıkıyoruz...
   Bu durum daha çok kurulup kaldırıldıktan sonra registry'deki çöplerini temizlemeyen programlar nedeniyle ortaya çıkar... }
 if (FHandler = nil) then Exit;
 
 { Bunu daha çok word ve excel kullanıyor... }
 if (FHandler.QueryInterface(IInitializeWithFile, FileInit) = 0) then begin
     FileInit.Initialize(StringToOleStr(FFileName), STGM_READ or STGM_SHARE_DENY_NONE);
     FKip := 'fileinit';
 end;
 { Bu noktada ELSE kullanılmamasının sebebi arabirimin her iki metodu da aynı anda destekleyebilecek şekilde tasarlanmasından kaynaklanıyor olabilir. Test edilmeli... UP }
 
 { Bunu da daha çok Acrobat Reader kullanıyor... }
 if (FHandler.QueryInterface(IInitializeWithStream, StreamInit) = 0) then begin
     { Hem dosya, hem de bellek akışı kullanmamızın tek bir nedeni var. Dosya akışı kullandığımızda işletim sistemi
       dosyayı (biz her ne kadar ShareDenyNone desek bile) bir sebepten dolayı işlem gören bir dosya olarak algılıyor,
       Bu sorunu aşmanın en basit yolu dosyayı belleğe kopyalayıp bellek kopyası üzerinde işlem yapmak. O nedenle dosyayı
       diskten okuduktan sonra doğruadn belleğe alıyoruz ve dosyayı kapatıyoruz. Bu sayede "Dosya başka bir işlem
       tarafından kullanılıyor..." gibi bir hata ile muhatap olmuyoruz... Aksi durumda dosyayı kendimiz açtığımız
       halde silememe gibi sorunlarla karşılaşırdık...
     }
     try
       try
         FS := TFileStream.Create(FFileName, fmOpenRead or fmShareDenyNone); // Dosyayı oku, paylaşım için herhangi bir kilit koyma, diğer programlar da o dosyaya erişebilsin...
         if (Assigned(FS) = TRUE) then begin
             TamponBellek:= TMemoryStream.Create;
             TamponBellek.Clear;
             TamponBellek.Size := FS.Size;
             TamponBellek.CopyFrom(FS, FS.Size);
         end;
       finally
         FS.Free; { Nihayetinde dosyayı kapat... Dosya diskte dursun, zaten yerini biliyoruz... }
       end;
     finally
       SA := TStreamAdapter.Create(TamponBellek, soOwned) as IStream;
       StreamInit.Initialize(SA, STGM_READ or STGM_SHARE_DENY_NONE);
 
       if  (FKip <> '')
       then FKip := format('%s + %s', [FKip, 'streaminit'])
       else FKip := 'streaminit';
     end;
 end;
 
 {
   Acrobat ve MS Office arasında IPreviewHandler arabiriminin kullanımıyla ilgili temel bir farklılık var.
   Ofis, yeni bir dosya ekleneceği zaman kendisini kapatıp yeni dosyayla birlikte belleğe yeniden yükleniyor.
   Acrobat ise aynı durumda kendisini kapatmıyor, sadece eski dosyayı kapatıp yeni dosyayı yüklüyor.
   O nedenle birisi fileinit'i diğeri ise FileStream'i kullanıyor. Yani normalde olması gereken ofisin yöntemi
   fakat StreamInit if bloğundaki işlemi bazı programlar kendi içinde değil, bizim yapmamızı yeğliyor. Temel fark aslında bu...
 }
 
 ARect := Rect( 0, 0, FControl.Width, FControl.Height);
 try
   FHandler.SetWindow(FControl.Handle, aRect);
   FHandler.SetRect(aRect);
 
   case FHandler.DoPreview of
         S_OK                         : Mesaj := ''; // 'The operation completed successfully.';
         E_PREVIEWHANDLER_DRM_FAIL    : Mesaj := 'Telif hakları gereği bu dosya dijital haklar yönetimi tarafından engellendi.'; //'Blocked by digital rights management.';
         E_PREVIEWHANDLER_NOAUTH      : Mesaj := 'Dosya izinleri engellendi.'; // 'Blocked by file permissions.';
         E_PREVIEWHANDLER_NOTFOUND    : Mesaj := 'Nesne bulunamadı.'; // 'Item was not found.';
         E_PREVIEWHANDLER_CORRUPT     : Mesaj := 'Önizleme yapan program düzgün yüklenmemiş veya ilgili fonksiyonu bozuk.'#13#10
                                               + 'Bu noktada ağ yöneticinizden yardım talep edebilirsiniz; Ağ yöneticiniz şunları yapmalıdır;'#13#10#13#10
                                               + ' 1) İlgili programı bilgisayarınızdan kaldırın (uninstall)'#13#10
                                               + ' 2) Kaldırma işlemi bittikten sonra Registry kayıtlarını temizleyen programlardan biriyle (CCleaner gibi) temizlik yapın'#13#10
                                               + ' 3) Kaldırdığınız programın eskiden kurulu olduğu Klasörleri silin (Genelde Program Files klasörünün altında olurlar)'#13#10
                                               + ' 4) 2. adımı tekrar yapın'#13#10
                                               + ' 5) Bilgisayarınızı yeniden başlatın'#13#10
                                               + ' 6) Kaldırdığınız programı sağlıklı bir şekilde yeniden kurun'#13#10
                                               ; // 'Item was corrupt.';
   end;
 
   if (Mesaj <> '') then begin
       ShowMessage('DoPreview : ' + Mesaj);
   end;
   FHandler._AddRef; { Referans sayımını tetiklemek ve yeni bir dosyanın yüklendiğini IPreviewHandler arabirimine bildirmek için bu şart... }
 except
   FHandler.Unload;
 end;
end;
 
procedure TdxPreviewHandler.Resize;
var
 aRect: TRect;
begin
 if (Assigned(FControl) = True) then begin
     if (FControl.Visible = TRUE) then begin
         if (Assigned(FHandler) = True) then begin
             try
               aRect := Rect( 0, 0, FControl.Width, FControl.Height);
               FHandler.SetWindow(FControl.Handle, aRect);
               FHandler.SetRect(aRect);
               // FHandler.DoPreview;      // gereksiz...
               // FHandler.SetFocus;       // gereksiz...
               if pos('streaminit', FKip) > 0 then FHandler._AddRef; // Şart ! .... Referans sayımını tetiklemek ve yeni bir dosyanın yüklendiğini IPreviewHandler arabirimine bildirmek için bu şart...
             except
               on E: Exception do ShowMessage('Hata: ' + E.Message);
             end;
         end;
     end;
 end;
end;
 
procedure TdxPreviewHandler.Preview;
begin
  PreviewFile;
end;
 
procedure TdxPreviewHandler.SetFileName(const Value: string);
begin
  if not FileExists(Value, True) then exit; // dosya yoksa doğrudan çık...
  if Trim(FFileName) = Trim(Value) then Exit; // Aynı şeyi arka arkaya yüklememek içindir...
  FFileName := Value;
  if  (GetClassID <> '')
  then PreviewFile  // Destekleniyorsa belgeyi ekranda gösterir.
  else Bosalt;      // Desteklenmiyorsa ekranı ve belleği temizler.
end;
 
end.
