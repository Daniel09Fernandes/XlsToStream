
function GeraXLSStream( Dataset: TFDQuery; MemDataset: TFDMemTable):TMemoryStream;
const
 limitfixo = 10000;
var
  iArchivo: TStringlist;
  Linea: string;
  i,ind,j: Integer;
  paginacao:integer;
  limit: integer;
  inicio:integer;
  f,k:integer;
begin

   //Daniel: -14/01/2021 O Download e realizado por esta função por causa da paginação do dataSet que prescisei fazer aqui,
   //após 10.000 registro o firedac da erro de falta de memoria, por isso divide o arquivo gerado

    Dataset.DisableControls;

     Dataset.last;
     Dataset.First;

     limit := limitfixo;
     inicio := 0;

    paginacao := trunc((Dataset.RecordCount / limitfixo) + 1);

    for i := 1 to paginacao do
    begin
      Result := TMemoryStream.Create;
      iArchivo := TStringlist.Create;
      iArchivo.Add('<HTML><BODY><table>');
      Linea := '<tr>';
      for ind := 0 to Dataset.FieldCount - 1 do
      begin
        if Dataset.Fields[ind].visible = True then
        begin
          Linea := Linea + '<td>' + Dataset.Fields[ind].DisplayLabel + '</td>';
        end;
      end;
      Linea := Linea + '</tr>';

      iArchivo.Add(Linea);


      while (inicio <> limit) and ( inicio < Dataset.recordcount ) do
      begin
        memDataset.Append;

        for f := 0 to Dataset.FieldCount - 1 do
        begin
            memDataset.Fields[f].AsString := Dataset.Fields[f].AsString;
        end;

        memDataset.Post;
        Dataset.Next;
        inc(inicio);
      end;

        if memDataset.RecordCount = 0 then
         exit;

      memDataset.Last;
      memDataset.First;

      while memDataset.RecNo <> memDataset.RecordCount do
      begin
        Linea := '<tr>';
        for k:= 0 to memDataset.FieldCount - 1 do
        begin
          if memDataset.Fields[k].Visible = True then
            if memDataset.Fields[k].DataType in [ftFloat, ftCurrency, ftBCD, ftFMTBcd, ftLargeint, ftSmallint, ftInteger, ftSingle] then
              Linea := Linea + '<td>' + memDataset.Fields[k].AsString + '</td>'
            else
              Linea := Linea + '<td style="width:auto;">' + memDataset.Fields[k].AsString + '</td>';
//            Linea := Linea + '<td>&nbsp;' + Dataset.Fields[i].AsString + '</td>';
        end;
        Linea := Linea + '</tr>';
        iArchivo.Add(Linea);
        Linea := '';
        memDataset.Next;
      end;
      Linea := ''; // limpa var memoria

      iArchivo.Add('</table></body></html>');
      iArchivo.SaveToStream(result);
     // iArchivo.savetofile('c:/bO'+FormatDateTime('ss',now)+'.txt');
      UniSession.SendStream(result,GeraNome+'-NF.xls');
      inicio := limit;
      limit := limit + limitfixo;
      memDataset.EmptyDataSet;
      iArchivo.Free;
      result.Free;
      Sleep(1000);

    end;
end;