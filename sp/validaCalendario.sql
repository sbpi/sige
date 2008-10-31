ALTER function dbo.ValidaCalendario(@calendario numeric(18)) returns varchar(2000) as
/**********************************************************************************
* Nome      : ValidaCalendario
* Finalidade: Valida as datas de um calend�rio
* Autor     : Alexandre Vinhadelli Papad�polis
* Data      :  31/10/2008, 14:43
* Par�metros:
*    @cliente       : indica o cliente. Se nulo, executa para todos os clientes.
*    @calendario    : indica o calend�rio do cliente. Se nulo, executa para todos os calend�rios
* Retorno: nulo - n�o foram encontrados erros para o calend�rio
*          string - lista de erros encontrados
***********************************************************************************/
begin
  Declare @w_existe   int;
  Declare @w_chave    numeric(18);
  Declare @w_tipo     numeric(18);
  Declare @w_data     datetime;
  Declare @w_sigla    varchar(255);
  Declare @w_nome     varchar(255);
  Declare @texto      varchar(2000);

  Set @texto = '';

  -- Verifica se o calend�rio tem in�cio/fim de ano/semestre letivo. Qtd deve ser igual a 1
  Declare c_ano_letivo cursor for
    select a.sq_particular_calendario, c.sigla, c.nome, count(*) qtd
      from escParticular_Calendario           a
           inner   join escCalendario_Cliente b on (a.sq_particular_calendario = b.sq_particular_calendario)
             inner join escTipo_Data          c on (b.sq_tipo_data             = c.sq_tipo_data)
     where a.sq_particular_calendario = @calendario
       and (c.sigla in ('IA','T1','I2','TA'))
    group by a.sq_particular_calendario, c.sigla, c.nome
    order by a.sq_particular_calendario, c.sigla, c.nome;

  -- Verifica se foi inserida alguma data que coincida com o calend�rio oficial
  Declare c_base cursor for
    select a.sq_particular_calendario, b.dt_ocorrencia, count(*) qtd
      from escParticular_Calendario           a
           inner   join escCalendario_Cliente b on (a.sq_particular_calendario = b.sq_particular_calendario)
             inner join escTipo_Data          c on (b.sq_tipo_data             = c.sq_tipo_data)
             inner join escCalendario_Base    d on (b.dt_ocorrencia            = d.dt_ocorrencia)
     where a.sq_particular_calendario = @calendario
    group by a.sq_particular_calendario, b.dt_ocorrencia
    order by a.sq_particular_calendario, b.dt_ocorrencia;

  -- Verifica se alguma data tem mais de uma ocorr�ncia com o mesmo tipo
  Declare c_duplicata cursor for
    select a.sq_particular_calendario, b.dt_ocorrencia, b.sq_tipo_data, count(*) qtd
      from escParticular_Calendario           a
           inner   join escCalendario_Cliente b on (a.sq_particular_calendario = b.sq_particular_calendario)
             inner join escTipo_Data          c on (b.sq_tipo_data             = c.sq_tipo_data)
     where a.sq_particular_calendario = @calendario
    group by a.sq_particular_calendario, b.dt_ocorrencia, b.sq_tipo_data
    having count(*) > 1
    order by a.sq_particular_calendario, b.dt_ocorrencia;

  -- Verifica se o calend�rio tem in�cio/fim de ano/semestre letivo. Qtd deve ser igual a 1
  Open c_ano_letivo
  Fetch Next from c_ano_letivo into @w_chave, @w_sigla, @w_nome, @w_existe;
  While @@fetch_status = 0 Begin
      If @w_existe = 0 Begin
         Set @texto = @texto + '<li>' + @w_nome + ' deve ser informado';
      End
      Fetch Next from c_ano_letivo into @w_chave, @w_nome, @w_sigla, @w_existe;
  End
  Close c_ano_letivo;
  Deallocate c_ano_letivo;

  -- Verifica se foi inserida alguma data que coincida com o calend�rio oficial
  Open c_base
  Fetch Next from c_base into @w_chave, @w_data, @w_existe;
  While @@fetch_status = 0 Begin
      If @w_existe > 0 Begin
         Set @texto = @texto + '<li>' + convert(varchar,@w_data,103) + ' consta do Calend�rio Oficial e deve ser removida deste calend�rio';
      End
      Fetch Next from c_base into @w_chave, @w_data, @w_existe;
  End
  Close c_base;
  Deallocate c_base;

  -- Verifica se alguma data tem mais de uma ocorr�ncia com o mesmo tipo
  Open c_duplicata
  Fetch Next from c_duplicata into @w_chave, @w_data, @w_tipo, @w_existe;
  While @@fetch_status = 0 Begin
      If @w_existe > 1 Begin
         Set @texto = @texto + '<li>' + convert(varchar,@w_data,103) + ' s� pode ser informada uma vez';
      End
      Fetch Next from c_duplicata into @w_chave, @w_data, @w_tipo, @w_existe;
  End
  Close c_duplicata;
  Deallocate c_duplicata;

  If @texto = '' Set @texto     = null;

  Return @texto;
end
