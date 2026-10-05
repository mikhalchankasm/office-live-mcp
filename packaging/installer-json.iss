// Small JSON reader for the CLI contract and the legacy state/marker hints.
// No agent configuration is interpreted here. Malformed/ambiguous input fails closed.
procedure JsonSpace(const S: String; var P: Integer);
begin
  while P <= Length(S) do begin
    if Pos(S[P], ' ' + #9 + #10 + #13) = 0 then Break;
    P := P + 1;
  end;
end;

function JsonString(const S: String; var P: Integer; var Value: String): Boolean;
var C: Char; N: Integer;
begin
  Result := False;
  Value := '';
  JsonSpace(S, P);
  if P > Length(S) then Exit;
  if S[P] <> '"' then Exit;
  P := P + 1;
  while P <= Length(S) do begin
    C := S[P]; P := P + 1;
    if C = '"' then begin Result := True; Exit; end;
    if Ord(C) < 32 then Exit;
    if C = '\' then begin
      if P > Length(S) then Exit;
      C := S[P]; P := P + 1;
      case C of
        '"', '\', '/': Value := Value + C;
        'b': Value := Value + #8;
        'f': Value := Value + #12;
        'n': Value := Value + #10;
        'r': Value := Value + #13;
        't': Value := Value + #9;
        'u': begin
          if P + 3 > Length(S) then Exit;
          N := StrToIntDef('$' + Copy(S, P, 4), -1);
          if N < 0 then Exit;
          Value := Value + Chr(N); P := P + 4;
        end;
      else Exit;
      end;
    end else Value := Value + C;
  end;
end;

function JsonValue(const S: String; var P: Integer; var Value: String): Boolean;
var Start, Depth: Integer; Ignored, Stack: String; C: Char;
begin
  Result := False;
  JsonSpace(S, P); Start := P;
  if P > Length(S) then Exit;
  if S[P] = '"' then begin
    Result := JsonString(S, P, Ignored);
    Value := Copy(S, Start, P - Start); Exit;
  end;
  if (S[P] = '[') or (S[P] = '{') then begin
    Stack := ''; Depth := 0;
    while P <= Length(S) do begin
      C := S[P];
      if C = '"' then begin
        if not JsonString(S, P, Ignored) then Exit;
      end else begin
        P := P + 1;
        if (C = '[') or (C = '{') then begin
          Stack := Stack + C; Depth := Depth + 1;
        end else if (C = ']') or (C = '}') then begin
          if Depth = 0 then Exit;
          if ((C = ']') and (Stack[Depth] <> '[')) or ((C = '}') and (Stack[Depth] <> '{')) then Exit;
          Delete(Stack, Depth, 1); Depth := Depth - 1;
          if Depth = 0 then begin
            Value := Copy(S, Start, P - Start); Result := True; Exit;
          end;
        end;
      end;
    end;
    Exit;
  end;
  while P <= Length(S) do begin
    if Pos(S[P], ',]} ' + #9 + #10 + #13) > 0 then Break;
    P := P + 1;
  end;
  Value := Copy(S, Start, P - Start);
  Result := (Value = 'true') or (Value = 'false') or (Value = 'null') or
    (StrToIntDef(Value, -1) >= 0);
end;

function JsonField(const S, Key: String; var Value: String): Boolean;
var P: Integer; Name, Item: String; Found: Boolean;
begin
  Result := False; Found := False; P := 1;
  JsonSpace(S, P);
  if P > Length(S) then Exit;
  if S[P] <> '{' then Exit;
  P := P + 1;
  while P <= Length(S) do begin
    if not JsonString(S, P, Name) then Exit;
    JsonSpace(S, P);
    if P > Length(S) then Exit;
    if S[P] <> ':' then Exit;
    P := P + 1;
    if not JsonValue(S, P, Item) then Exit;
    if Name = Key then begin
      if Found then Exit;
      Value := Item; Found := True;
    end;
    JsonSpace(S, P);
    if P > Length(S) then Exit;
    if S[P] = '}' then begin
      P := P + 1; JsonSpace(S, P);
      Result := Found and (P > Length(S)); Exit;
    end;
    if S[P] <> ',' then Exit;
    P := P + 1;
  end;
end;

function JsonText(const S, Key: String): String;
var Raw: String; P: Integer;
begin
  Result := '';
  if not JsonField(S, Key, Raw) then Exit;
  P := 1;
  if not JsonString(Raw, P, Result) then Result := '';
end;

function JsonTrue(const S, Key: String): Boolean;
var Raw: String;
begin
  Result := JsonField(S, Key, Raw) and (Raw = 'true');
end;

function JsonArray(const S: String; var Items: TArrayOfString): Boolean;
var P, N: Integer; Item: String;
begin
  Result := False; P := 1; SetArrayLength(Items, 0);
  JsonSpace(S, P);
  if P > Length(S) then Exit;
  if S[P] <> '[' then Exit;
  P := P + 1; JsonSpace(S, P);
  if P > Length(S) then Exit;
  if S[P] <> ']' then begin
    while P <= Length(S) do begin
      if not JsonValue(S, P, Item) then Exit;
      N := GetArrayLength(Items); SetArrayLength(Items, N + 1); Items[N] := Item;
      JsonSpace(S, P);
      if P > Length(S) then Exit;
      if S[P] = ']' then Break;
      if S[P] <> ',' then Exit;
      P := P + 1;
    end;
  end;
  P := P + 1; JsonSpace(S, P); Result := P > Length(S);
end;

function JsonProblems(const S: String): String;
var Raw, Item: String; Items: TArrayOfString; I, P: Integer;
begin
  Result := '';
  if not JsonField(S, 'problems', Raw) then Exit;
  if not JsonArray(Raw, Items) then Exit;
  for I := 0 to GetArrayLength(Items) - 1 do begin
    P := 1;
    if JsonString(Items[I], P, Item) then Result := Result + Item + #13#10;
  end;
end;
