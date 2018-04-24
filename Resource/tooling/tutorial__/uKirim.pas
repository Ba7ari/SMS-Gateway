unit uKirim;

interface

procedure KirimSMS(smsc, Tujuan, Isi: string);

const
	sOK = 'OK';
	sERROR = 'ERROR';
	ProXLsmsc   ='62818445009';
	Simpatismsc ='6281100000';
	Mentarismsc ='62816124';
	IM3smsc     ='62855000000';

var
	lkirim : integer;

implementation

Uses uTerimaSMS;

function text2PDU(text:string):string;
var   PDU : string;
      geser,panjang,tmp,tmp2,tmp3,n:byte;
begin
      PDU :='';
      panjang :=length(text);
      PDU :=PDU+inttohex(panjang,2);

      geser :=0;
      for n :=1 to panjang-1 do
      begin
          tmp2 :=ord(text[n]);
          if geser<>0 then tmp2 :=tmp2 shr geser;
          tmp :=ord(text[n+1]);
          if geser=7 then
          begin
                geser :=0;
          end else
          begin
                tmp3 :=8-(geser+1);
                if tmp3<>0 then tmp:=tmp shl tmp3;
                PDU :=PDU+inttohex((tmp or tmp2),2);
                inc(geser);
          end;
      end;
      if geser<7 then
      begin
          tmp2:=ord(text[panjang]);
          if(geser<>0)then tmp2:=tmp2 shr geser;
          PDU:=PDU+inttohex(tmp2,2);
      end;
      result:=PDU;
end;

function ConvertText(smsc,tipe,ref,tujuan,bentuk,skema,validitas,isi:string):string;
var
   PDU,tmp :string;
   p,
   i :byte;
begin
   PDU := '';
   If length(smsc)=0 then begin result :=''; exit; end;
   if length(tipe)=0 then tipe :='11';
   if length(ref)=0 then ref :='00';
   if length(bentuk)=0 then bentuk :='00';
   if length(skema)=0 then skema :='00';
   if length(validitas)=0 then validitas :='FF';
   If length(isi)=0 then begin result :=''; exit; end;

   if smsc[1]='0' then tmp :='81'+smsc else tmp :='91'+smsc;
   if(length(tmp)mod 2)<>0 then tmp :=tmp+'F';
   p := length(tmp);
   PDU :=PDU + inttohex(p div 2,2) + tmp[1] + tmp[2];
   for i:= 2 to length(tmp)div 2 do
   begin
        PDU :=PDU+tmp[i*2];
        PDU :=PDU+tmp[(i*2)-1];
   end;
   PDU :=PDU+tipe;
   PDU :=PDU+ref;
   if tujuan[1]='+' then tujuan:=copy(tujuan,2,length(tujuan)-1);
   PDU :=PDU+inttohex(length(tujuan),2);
   if(length(tujuan)mod 2)<>0 then tujuan:=tujuan+'F';
   if tujuan[1]='0' then PDU :=PDU+'81' else begin
     PDU :=PDU+'91';
   end;
   for i :=1 to length(tujuan)div 2 do
   begin
        PDU :=PDU+tujuan[i*2];
        PDU :=PDU+tujuan[(i*2)-1];
   end;
   PDU :=PDU+bentuk;
   PDU :=PDU+skema;
   PDU :=PDU+validitas;
   tmp := Text2PDU(isi);
   PDU :=PDU+tmp;
   i := length(tmp);
   lkirim := (length(PDU) - p) div 2;
   result :=PDU;
end;

function SendGetData(teks: string; tOK: string): string;
var waktu  : TDateTime;
    buffer : string;
begin
     waktu := now;
     comm.Output := teks;
     sleep(500);
     buffer := '';
     repeat
           buffer := buffer + comm.Input;
     until (pos(tOK, buffer) > 0) or (pos(sERROR, buffer) > 0)
           or (SecondsBetween(waktu,now) > 60);
     result := buffer;
end;

procedure KirimSMS;
var
   PDU :string;
begin
        PDU := ConvertText(IM3smsc,'','',Tujuan,'','','',isi);
        SendGetData('AT+CMGS=' + inttostr(lkirim) + #13, '>');
        SendGetData(PDU + #$1A, sOK);
end;

begin

end.