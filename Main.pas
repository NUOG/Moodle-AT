unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, Grids, frxClass, frxExportPDF,
  frxExportRTF, Buttons;

type
  TMoodleAD = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    StatusBar1: TStatusBar;
    LabeledEdit1: TLabeledEdit;
    Button1: TButton;
    LabeledEdit2: TLabeledEdit;
    Button2: TButton;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    StringGrid1: TStringGrid;
    Button4: TButton;
    LabeledEdit3: TLabeledEdit;
    Button5: TButton;
    LabeledEdit4: TLabeledEdit;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    DateTimePicker1: TDateTimePicker;
    ComboBox5: TComboBox;
    ComboBox6: TComboBox;
    Edit1: TEdit;
    Button6: TButton;
    Button7: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    frxPDFExport1: TfrxPDFExport;
    frxReport1: TfrxReport;
    frxUserDataSet1: TfrxUserDataSet;
    Button8: TButton;
    StringGrid2: TStringGrid;
    Button9: TButton;
    frxReport2: TfrxReport;
    frxPDFExport2: TfrxPDFExport;
    frxUserDataSet2: TfrxUserDataSet;
    Label9: TLabel;
    Label10: TLabel;
    ComboBox7: TComboBox;
    ComboBox8: TComboBox;
    Label11: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    ComboBox9: TComboBox;
    Label12: TLabel;
    ComboBox10: TComboBox;
    Label13: TLabel;
    frxRTFExport1: TfrxRTFExport;
    export_to_word: TButton;
    Image1: TImage;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure frxReport1GetValue(const VarName: String;
      var Value: Variant);
    procedure Button8Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure frxReport2GetValue(const VarName: String;
      var Value: Variant);
    procedure FormShow(Sender: TObject);
    procedure export_to_wordClick(Sender: TObject);
    procedure Image1DblClick(Sender: TObject);
  private
    stl, st2, st3, st4, st5, st6, st7, st8: TStringList;
    { Private declarations }
  public
    procedure csv2grid;
    procedure csv2grid2;
    procedure GenPassword;
    procedure FillReport;
    procedure AutoSizeGrid(Grid: TStringGrid);
    { Public declarations }
  end;

var
  MoodleAD: TMoodleAD;

implementation

{$R *.dfm}
procedure TMoodleAD.AutoSizeGrid(Grid: TStringGrid);
const
  ColWidthMin = 10;
var
  C, R, W, ColWidthMax: integer;
begin
  for C := 0 to Grid.ColCount - 1 do begin
    ColWidthMax := ColWidthMin;
    for R := 0 to (Grid.RowCount - 1) do begin
      W := Grid.Canvas.TextWidth(Grid.Cells[C, R]);
      if W > ColWidthMax then
        ColWidthMax := W+10;
    end;
    Grid.ColWidths[C] := ColWidthMax + 5;
  end;
end;

function RamdomPassword () : String;
const
  intMAX_PW_LEN = 8;
var
  i: Byte;
  s: string;
begin
  Randomize;
  {if you want to use the 'A..Z' characters}
  if true then
    s := 'ABCDEFGHJKMNPQRSTUVWXYZ'
  else
    s := '';

  {if you want to use the 'a..z' characters}
  if true then
    s := s + 'abcdefghjkmnopqrstuvwxyz';

  {if you want to use the '0..9' characters}
  if true then
    s := s + '0123456789';

  if true then
    s := s + '#$%';

  if s = '' then exit;

  Result := '';
  for i := 0 to intMAX_PW_LEN-1 do
    Result := Result + s[Random(Length(s)-1)+1];
end;

function RepalceComma (s : string) : String;
var
 i : integer;
Begin
result := '';
 for i := 1 to length (s) do
  if s[i] <> '.' then
   result := result + s[i]
  else
   result := result + ',';
End;

function RepalceApostrof (s : string) : String;
var
 i : integer;
Begin
result := '';
 for i := 1 to length (s) do
  if s[i] <> '''' then
   result := result + s[i]
  else
   result := result + '&#39;';
End;

function CutSymbol (s : string; sym: string) : String;
var
 i : integer;
Begin
result := '';
 for i := 1 to length (s) do
  if s[i] <> sym then
   result := result + s[i];
End;

function cutspace (s : string) : String;
var
 i : integer;
Begin
result := '';
 for i := 1 to length (s) do
  if s[i] <> ' ' then
   result := result + s[i];
End;

function Translit(s: string): string;
const
  ukr: string = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюяіїєАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЬЫЪЭЮЯІЇЄ';
  lat: array[1..72] of string = ('a', 'b', 'v', 'g', 'd', 'e', 'yo', 'zh', 'z', 'y', 'yi', 'k', 'l', 'm', 'n', 'o', 'p', 'r', 's', 't', 'u', 'f', 'kh', 'ts', 'ch', 'sh', 'shch', '''', 'y', '''', 'e', 'yu', 'ya', 'i', 'yi', 'ye', 'A', 'B', 'V', 'G', 'D', 'E', 'Yo', 'Zh', 'Z', 'Y', 'Yi', 'K', 'L', 'M', 'N', 'O', 'P', 'R', 'S', 'T', 'U', 'F', 'Kh', 'Ts', 'Ch', 'Sh', 'Shch', '''', 'Y', '''', 'E', 'Yu', 'Ya', 'I', 'Yi', 'Ye');
var
  p, i, l: integer;
begin
  Result := '';
  l := Length(s);
  for i := 1 to l do
  begin
    p := Pos(s[i], ukr);
    if p<1 then Result := Result + s[i] else Result := Result + lat[p];
  end;
end;

function PointToECTS (point: String) : String;
begin
 Result := '';
 if (StrToFloat (point) >= 0) and (StrToFloat (point) <= 34) then
  Result := 'F';
 if (MoodleAD.ComboBox10.Text = '1') then
  begin
   if (StrToFloat (point) >= 35) and (StrToFloat (point) <= 49) then
    Result := 'FX';
   if (StrToFloat (point) >= 50) and (StrToFloat (point) <= 66) then
    Result := 'E';
  end
 else
  begin
   if (StrToFloat (point) >= 35) and (StrToFloat (point) <= 59) then
    Result := 'FX';
   if (StrToFloat (point) >= 60) and (StrToFloat (point) <= 66) then
    Result := 'E';
  end;
 if (StrToFloat (point) >= 67) and (StrToFloat (point) <= 74) then
  Result := 'D';
 if (StrToFloat (point) >= 75) and (StrToFloat (point) <= 81) then
  Result := 'C';
 if (StrToFloat (point) >= 82) and (StrToFloat (point) <= 89) then
  Result := 'B';
 if (StrToFloat (point) >= 90) and (StrToFloat (point) <= 100) then
  Result := 'A';
end;

function PointTo5 (point: String) : String;
begin
if (MoodleAD.ComboBox10.Text = '1') then
begin
 Result := '';
 if point = '' then
  Result := 'н/я';
 if (StrToFloat (point) >= 0) and (StrToFloat (point) <= 49) then
  Result := 'незадовільно';
 if (StrToFloat (point) >= 50) and (StrToFloat (point) <= 74) then
  Result := 'задовільно';
 if (StrToFloat (point) >= 75) and (StrToFloat (point) <= 89) then
  Result := 'добре';
 if (StrToFloat (point) >= 90) and (StrToFloat (point) <= 100) then
  Result := 'відмінно';
end
else
begin
 Result := '';
 if point = '' then
  Result := 'н/я';
 if (StrToFloat (point) >= 0) and (StrToFloat (point) <= 59) then
  Result := 'незадовільно';
 if (StrToFloat (point) >= 60) and (StrToFloat (point) <= 74) then
  Result := 'задовільно';
 if (StrToFloat (point) >= 75) and (StrToFloat (point) <= 89) then
  Result := 'добре';
 if (StrToFloat (point) >= 90) and (StrToFloat (point) <= 100) then
  Result := 'відмінно';
end;
end;

procedure TMoodleAD.csv2grid;
Var
 fn: TextFile;
 s, s1: String;
begin
StringGrid1.RowCount :=1;
if FileExists(LabeledEdit1.Text) then
 begin
  AssignFile (fn,LabeledEdit1.Text);
  Reset (fn);
  readln (fn, s);
  while not eof(fn) do
   begin
    readln (fn, s);
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid1.Cells [0,StringGrid1.RowCount-1] := cutspace(s1);   //group

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid1.Cells [1,StringGrid1.RowCount-1] := cutspace(s1);   //nomer zalikovoi

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid1.Cells [2,StringGrid1.RowCount-1] := cutspace(s1);   //prizvyshche

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid1.Cells [3,StringGrid1.RowCount-1] := cutspace(s1);   //name

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid1.Cells [4,StringGrid1.RowCount-1] := cutspace(s1);   //father name

    s := copy (s,pos(';',s)+1,length(s));
    StringGrid1.Cells [5,StringGrid1.RowCount-1] := cutspace(s);    //point =)


    StringGrid1.RowCount := StringGrid1.RowCount + 1;
   end;
  StringGrid1.RowCount := StringGrid1.RowCount - 1;
  CloseFile (fn);
 end;
end;

procedure TMoodleAD.csv2grid2;
Var
 fn, fn2: TextFile;
 s, s1: String;
 ball, koef: real;
 ball1, ball2, k, i : Integer;
begin
for i := 0 to StringGrid2.RowCount - 1 
do StringGrid2.Rows[i].Clear;
StringGrid2.RowCount := 1;
if FileExists(LabeledEdit4.Text) and FileExists(LabeledEdit3.Text) then
 begin
  AssignFile (fn,LabeledEdit4.Text);
  Reset (fn);
  readln (fn, s);
  while not eof(fn) do
   begin
    readln (fn, s);
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid2.Cells [0,StringGrid2.RowCount-1] := s1;   //group

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid2.Cells [1,StringGrid2.RowCount-1] := s1;   //nomer zalikovoi

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid2.Cells [2,StringGrid2.RowCount-1] := cutspace(s1);   //prizvyshche

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid2.Cells [3,StringGrid2.RowCount-1] := cutspace(s1);   //name

    s := copy (s,pos(';',s)+1,length(s));
    s1 := copy (s,1,pos(';',s)-1);
    StringGrid2.Cells [4,StringGrid2.RowCount-1] := cutspace(s1);   //father name

    s := copy (s,pos(';',s)+1,length(s));
    StringGrid2.Cells [5,StringGrid2.RowCount-1] := s;    //point =)

    // Заповнюємо результати тестування якщо студент тестувався
    ball := 0;
    ball1 := 0;
    ball2 := 0;
    AssignFile (fn2,LabeledEdit3.Text);
      Reset (fn2);
      readln (fn2, s);
      while not eof(fn2) do
       begin
        readln (fn2, s);
        s1 := copy (s,1,pos(';',s)-1);
         if s1 = StringGrid2.Cells [2,StringGrid2.RowCount-1]+ ' '+
           StringGrid2.Cells [3,StringGrid2.RowCount-1]+ ' '+
           StringGrid2.Cells [4,StringGrid2.RowCount-1] then
            begin
              s := copy (s,pos(';',s)+1,length(s));
              s := copy (s,pos(';',s)+1,length(s));
              s := copy (s,pos(';',s)+1,length(s));
              s := copy (s,pos(';',s)+1,length(s));
              StringGrid2.Cells [6,StringGrid2.RowCount-1] := FloatToStr(Round(StrToFloat(RepalceComma(s))));   //ocinka za ekzamen

               //для тих хто здавав ыспит рахуэмо
                ball1 := StrToInt(StringGrid2.Cells [5,StringGrid2.RowCount-1]); //za semestr
                ball2 := StrToInt(StringGrid2.Cells [6,StringGrid2.RowCount-1]); //za test
                koef := StrToFloat(RepalceComma(Edit1.Text));
                if ball2 <> 0 then
                  begin
                    ball := Round((ball1+ball2*koef)/2);
                    if ball > 100 then ball := 100;
                  end;
                //if (ball2 = 0) and (ball1 >= 90) then
                //  ball := ball1;
                StringGrid2.Cells [7,StringGrid2.RowCount-1] := FloatToStr(ball);
                StringGrid2.Cells [8,StringGrid2.RowCount-1] := PointToECTS(StringGrid2.Cells [7,StringGrid2.RowCount-1]);
                StringGrid2.Cells [9,StringGrid2.RowCount-1] := PointTo5(StringGrid2.Cells [7,StringGrid2.RowCount-1]);
            end;
            // тут рахуємо для відмінників
           // if (ball2 = 0) and (ball1 >= 90) then
                ball := ball1;
           if (StrToFloat(StringGrid2.Cells [5,StringGrid2.RowCount-1]) >= 90) and (ball2 = 0) then
            begin
             ball := ball1;
             StringGrid2.Cells [7,StringGrid2.RowCount-1] := StringGrid2.Cells [5,StringGrid2.RowCount-1];
             StringGrid2.Cells [8,StringGrid2.RowCount-1] := PointToECTS(StringGrid2.Cells [5,StringGrid2.RowCount-1]);
             StringGrid2.Cells [9,StringGrid2.RowCount-1] := PointTo5(StringGrid2.Cells [5,StringGrid2.RowCount-1]);
            end;
        end;
      CloseFile (fn2);
      if StringGrid2.Cells [7,StringGrid2.RowCount-1] = '' then
        StringGrid2.Cells [9,StringGrid2.RowCount-1] := 'н/я';
    StringGrid2.RowCount := StringGrid2.RowCount + 1;
   end;
  StringGrid2.RowCount := StringGrid2.RowCount - 1;
  CloseFile (fn);
 end;
end;

procedure TMoodleAD.GenPassword;
Var
 i: Integer;
 fn: TextFile;
 s : String;
 login, password : String;
begin
 AssignFile (fn,LabeledEdit2.Text);
 Rewrite (fn);
 { username, password, firstname, lastname, email, lang, idnumber, maildisplay, course1, group1, type1
   jonest, verysecret, Вася, Коваль, koval@someplace.edu, uk, 3663737, 1, Intro101, Section 1, 1 }

 WriteLn (fn, 'username;password;firstname;lastname;email;lang');

  stl := TStringList.Create;
  st2 := TStringList.Create;
  st3 := TStringList.Create;

 for i := 0 to StringGrid1.RowCount - 1 do
  Begin
   login := LowerCase (Translit(StringGrid1.Cells [1,i]));
   password := RamdomPassword;
   WriteLn (fn, login+';'+
   password+';'+
   StringGrid1.Cells [3,i] + ' ' +StringGrid1.Cells [4,i]+';' +
   StringGrid1.Cells [2,i]+';'+
   LowerCase (Translit(StringGrid1.Cells [1,i]))+'@'+
   LowerCase (Translit(StringGrid1.Cells [0,i]))+'.nung.edu.ua;'+'uk'
   );
   stl.Add(StringGrid1.Cells [2,i]+' '+StringGrid1.Cells [3,i]+' '+StringGrid1.Cells [4,i]);
   st2.Add(login);
   st3.Add(password);
  End;

  frxUserDataSet2.RangeEnd := reCount;
  frxUserDataSet2.RangeEndCount := stl.Count;

 CloseFile (fn);
 s := StringGrid1.Cells [0,0];
 frxReport2.Variables['group'] := '''' + s + '''';
 if StringGrid1.RowCount > 2 then
  //frxReport2.ShowReport();
   s := LabeledEdit1.Text;
   frxPDFExport2.FileName := ExtractFilePath(s)+ copy(ExtractFileName(s),1,length(ExtractFileName(s))-length(ExtractFileExt(s)))+'.pass.pdf';
   frxReport2.PrepareReport();
   frxPDFExport2.ShowDialog:=False;
   frxReport2.Export(frxPDFExport2);

end;

procedure TMoodleAD.Button1Click(Sender: TObject);
var
S:String;
begin
  OpenDialog1.Execute;
  S:=OpenDialog1.FileName;
  LabeledEdit1.Text:=S;
  LabeledEdit4.Text:=S;
  LabeledEdit2.Text:= ExtractFilePath(s)+ copy(ExtractFileName(s),1,length(ExtractFileName(s))-length(ExtractFileExt(s)))+'.pass';
end;

procedure TMoodleAD.Button5Click(Sender: TObject);
var
S:String;
begin
OpenDialog1.Execute;
S:=OpenDialog1.FileName;
LabeledEdit4.Text:=S;
end;

procedure TMoodleAD.Button4Click(Sender: TObject);
var
S:String;
begin
OpenDialog1.Execute;
S:=OpenDialog1.FileName;
LabeledEdit3.Text:=S;
end;

procedure TMoodleAD.Button2Click(Sender: TObject);
begin
SaveDialog1.Execute;
LabeledEdit2.Text:=SaveDialog1.FileName;
end;

procedure TMoodleAD.Button3Click(Sender: TObject);
begin
GenPassword;
end;

procedure TMoodleAD.FillReport;
var
  s, summa : String;
  i, k : Integer;
begin
s := ComboBox1.Text;
frxReport1.Variables['nomervid'] := '''' + s + '''';
s := RepalceApostrof(ComboBox3.Text);
frxReport1.Variables['fakultet'] := '''' + s + '''';
s := RepalceApostrof(ComboBox9.Text);
frxReport1.Variables['naprjam'] := '''' + s + '''';
s := ComboBox10.Text;
frxReport1.Variables['course'] := '''' + s + '''';
s := ComboBox4.Text;
frxReport1.Variables['group'] := '''' + s + '''';
s := ComboBox2.Text;
frxReport1.Variables['semestr'] := '''' + s + '''';
s := RepalceApostrof(ComboBox5.Text);
frxReport1.Variables['discipline'] := '''' + s + '''';
s := RepalceApostrof(ComboBox8.Text);
frxReport1.Variables['vyd_kontrolu'] := '''' + s + '''';
s := RepalceApostrof(ComboBox6.Text);
frxReport1.Variables['teacher'] := '''' + s + '''';
s := RepalceApostrof(ComboBox7.Text);
frxReport1.Variables['teacher2'] := '''' + s + '''';
s := DateToStr(DateTimePicker1.Date);
frxReport1.Variables['dataex'] := '''' + s + '''';

stl := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо порядковий номер
  stl.Add(IntToStr(i));
 end;
st2 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо заліковки
  st2.Add(StringGrid2.Cells [1,i-1]);
 end;
st3 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо ПІП
  st3.Add(StringGrid2.Cells [2,i-1]+' '+StringGrid2.Cells [3,i-1]+' '+StringGrid2.Cells [4,i-1]);
 end;
st4 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо іспитову оцінку
  st4.Add(StringGrid2.Cells [6,i-1]);
 end;
st5 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо семестрову оцінку
  st5.Add(StringGrid2.Cells [5,i-1]);
 end;
st6 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо підсумкову 100 балів оцінку
  st6.Add(StringGrid2.Cells [7,i-1]);
 end;
st7 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо підсумкову ECTS оцінку
  st7.Add(StringGrid2.Cells [8,i-1]);
 end;
st8 := TStringList.Create;
for i := 1 to StringGrid2.RowCount do
 begin                                //Заповнюємо підсумкову 4 бальну оцінку
  st8.Add(StringGrid2.Cells [9,i-1]);
 end;

  frxUserDataSet1.RangeEnd := reCount;
  frxUserDataSet1.RangeEndCount := stl.Count;

//  summa := 'Всього: <b>'+IntToStr(StringGrid2.RowCount)+
//           '</b>, з них: відмінно ';
  frxReport1.Variables['st_count'] := '''' +
     IntToStr(StringGrid2.RowCount) + '''';
  k := 0;                                   //перевіряємо відмінників
  for i := 0 to StringGrid2.RowCount do
   begin
    if StringGrid2.Cells [7,i] = '' then
     s:= '0'
    else
     s := StringGrid2.Cells [7,i];
    if StrToInt(s) >= 90 then
     k := k + 1;
   end;
//   summa := summa + '<b>'+IntToStr(k)+'</b>, добре ';
  frxReport1.Variables['st_1'] := '''' +
     IntToStr(k) + '''';

  k := 0;                                   //перевіряємо студентів з доброю оцінкою
  for i := 0 to StringGrid2.RowCount do
   begin
    if StringGrid2.Cells [7,i] = '' then
     s:= '0'
    else
     s := StringGrid2.Cells [7,i];
    if (StrToInt(s) <= 89) and (StrToInt(s) >= 75) then
     k := k + 1;
   end;
//   summa := summa + '<b>'+IntToStr(k)+'</b>, задовільно ';
  frxReport1.Variables['st_2'] := '''' +
     IntToStr(k) + '''';

  k := 0;                                   //перевіряємо студентів з задовільною оцінкою
  for i := 0 to StringGrid2.RowCount do
   begin
    if StringGrid2.Cells [7,i] = '' then
     s:= '0'
    else
     s := StringGrid2.Cells [7,i];
    if (MoodleAD.ComboBox10.Text = '1') then
    begin
     if (StrToInt(s) <= 74) and (StrToInt(s) >= 50) then
      k := k + 1;
    end
    else
    begin
     if (StrToInt(s) <= 74) and (StrToInt(s) >= 60) then
      k := k + 1;
    end;

   end;
//   summa := summa + '<b>'+IntToStr(k)+'</b>,<br> незадовільно ';
  frxReport1.Variables['st_3'] := '''' +
     IntToStr(k) + '''';

  k := 0;                                   //перевіряємо студентів з не задовільною оцінкою
  for i := 0 to StringGrid2.RowCount do
   begin
    if StringGrid2.Cells [7,i] = '' then
     s:= '0'
    else
     s := StringGrid2.Cells [7,i];
    if (MoodleAD.ComboBox10.Text = '1') then
    begin
    if (StrToInt(s) <= 49) and (StrToInt(s) >= 0) and (StringGrid2.Cells [7,i] <> '') then
     k := k + 1;
    end
    else
    begin
    if (StrToInt(s) <= 59) and (StrToInt(s) >= 0) and (StringGrid2.Cells [7,i] <> '') then
     k := k + 1;
    end

   end;
//   summa := summa + '<b>'+IntToStr(k)+'</b>';
  frxReport1.Variables['st_4'] := '''' +
     IntToStr(k) + '''';


  k := 0;                                   //перевіряємо студентів з не задовільною оцінкою
  for i := 0 to StringGrid2.RowCount-1 do
   begin
    if StringGrid2.Cells [7,i] = '' then
     k := k + 1;
   end;
   if k > 0 then
//    summa := summa +', неявка '+ '<b>'+IntToStr(k)+'</b>';
  frxReport1.Variables['st_5'] := '''' +
     IntToStr(k) + ''''
   else
  frxReport1.Variables['st_5'] := '''' +
     ' ' + '''';

  frxReport1.Variables['st_6'] := '''' +
     ' ' + '''';
  frxReport1.Variables['st_7'] := '''' +
     ' ' + '''';


//  frxReport1.Variables['pidsumok'] := '''' + summa + '''';
// [st_coutn]

end;


procedure TMoodleAD.Button6Click(Sender: TObject);
begin
if StringGrid2.RowCount >= 1 then
 begin
  FillReport;
  frxReport1.ShowReport();
 end;
end;

procedure TMoodleAD.Button8Click(Sender: TObject);
begin
csv2grid;
AutoSizeGrid(StringGrid1);
end;

procedure TMoodleAD.Button7Click(Sender: TObject);
begin

if StringGrid2.RowCount >= 1 then
 begin
  FillReport;
  // frxReport1.ShowReport();
  // frxReport1.Preview := nil;
  SaveDialog1.FileName := Translit(ComboBox4.Text)+'-'+Translit(ComboBox5.Text);
  SaveDialog1.DefaultExt := 'pdf';
  SaveDialog1.Filter := 'Екзаменаційна відомість|*.pdf';
   if SaveDialog1.Execute then
    begin
     frxPDFExport1.FileName := SaveDialog1.FileName;
     frxReport1.PrepareReport();
     frxPDFExport1.ShowDialog:=False;
     frxReport1.Export(frxPDFExport1);
    end;
 end;
end;

procedure TMoodleAD.Button9Click(Sender: TObject);
begin
csv2grid2;
AutoSizeGrid(StringGrid2);
end;

procedure TMoodleAD.frxReport1GetValue(const VarName: String;
  var Value: Variant);
begin
  if CompareText(VarName, 'number') = 0 then
    Value := stl[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'zalikova') = 0 then
    Value := st2[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'student-name') = 0 then
    Value := st3[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'point1') = 0 then
    Value := st4[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'point2') = 0 then
    Value := st5[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'point3') = 0 then
    Value := st6[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'point4') = 0 then
    Value := st7[frxUserDataSet1.RecNo];
  if CompareText(VarName, 'point5') = 0 then
    Value := st8[frxUserDataSet1.RecNo];


end;

procedure TMoodleAD.frxReport2GetValue(const VarName: String;
  var Value: Variant);
begin
  if CompareText(VarName, 'student-name') = 0 then
    Value := stl[frxUserDataSet2.RecNo];
  if CompareText(VarName, 'login') = 0 then
    Value := st2[frxUserDataSet2.RecNo];
  if CompareText(VarName, 'password') = 0 then
    Value := st3[frxUserDataSet2.RecNo];
end;

procedure TMoodleAD.FormShow(Sender: TObject);
begin
DateTimePicker1.Date := now;
end;

procedure TMoodleAD.export_to_wordClick(Sender: TObject);
begin
if StringGrid2.RowCount >= 1 then
 begin
  FillReport;
  // frxReport1.ShowReport();
  // frxReport1.Preview := nil;
  SaveDialog1.FileName := Translit(ComboBox4.Text)+'-'+Translit(ComboBox5.Text);
  SaveDialog1.DefaultExt := 'doc';
  SaveDialog1.Filter := 'Екзаменаційна відомість|*.doc';
   if SaveDialog1.Execute then
    begin
     frxRTFExport1.FileName := SaveDialog1.FileName;
     frxReport1.PrepareReport();
     frxRTFExport1.ShowDialog:=False;
     frxReport1.Export(frxRTFExport1);
    end;
 end;
end;

procedure TMoodleAD.Image1DblClick(Sender: TObject);
begin
if StringGrid2.RowCount >= 1 then
 begin
  FillReport;
  // frxReport1.ShowReport();
  // frxReport1.Preview := nil;
  SaveDialog1.FileName := Translit(ComboBox4.Text)+'-'+Translit(ComboBox5.Text);
  SaveDialog1.DefaultExt := 'doc';
  SaveDialog1.Filter := 'Екзаменаційна відомість|*.doc';
   if SaveDialog1.Execute then
    begin
     frxRTFExport1.FileName := SaveDialog1.FileName;
     frxReport1.PrepareReport();
     frxRTFExport1.ShowDialog:=False;
     frxReport1.Export(frxRTFExport1);
    end;
 end;
end;

end.
