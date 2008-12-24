create or replace function YEAR(VAL IN DATE) return number is
  Result NUMBER;
begin
  Result := to_char(val,'yyyy');
  return(Result);
end YEAR;
/
