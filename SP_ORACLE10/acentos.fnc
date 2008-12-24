CREATE OR REPLACE FUNCTION ACENTOS ( Valor IN VARCHAR2) RETURN  VARCHAR2 IS
   nome varchar2(8000) := trim(Valor);
BEGIN
   nome := replace(replace(translate(nome,'ΐΒΑΓΚΙΝΤΥΣΪάΗΰβαγκιντυσϊόη','AAAAEEIOOOUUCaaaaeeiooouuc'),'&','e'),'-','- ');

   RETURN upper(nome) ;
END;
/
