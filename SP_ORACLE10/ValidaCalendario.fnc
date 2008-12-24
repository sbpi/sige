create or replace function ValidaCalendario(calendario in number, ano in number) return varchar2 is
/**********************************************************************************
* Nome      : ValidaCalendario
* Finalidade: Valida as datas de um calend�rio
* Autor     : Alexandre Vinhadelli Papad�polis
* Data      :  31/10/2008, 14:43
* Par�metros:
*    cliente       : indica o cliente. Se nulo, executa para todos os clientes.
*    calendario    : indica o calend�rio do cliente. Se nulo, executa para todos os calend�rios
* Retorno: nulo - n�o foram encontrados erros para o calend�rio
*          string - lista de erros encontrados
***********************************************************************************/
  texto      varchar2(2000) := '';

  cursor c_ano_letivo is
    select c.sigla, c.nome, count(*) qtd
      from Tipo_Data c
           left join (select a.sq_particular_calendario, b.sq_tipo_data
                        from Calendario_Cliente               b
                             inner join Particular_Calendario a on (a.sq_particular_calendario = b.sq_particular_calendario)
                       where a.sq_particular_calendario = calendario
                       and   year(b.dt_ocorrencia) = ano
                     )  d on (d.sq_tipo_data             = c.sq_tipo_data)
     where (c.sigla in ('IA','T1','I2','TA'))
       and d.sq_tipo_data is null
    group by c.sigla, c.nome
    order by c.sigla, c.nome;

  -- Impede duplica��o nas datas de in�cio/fim de ano/semestre letivo
  cursor c_ano_letivo_dup is
    select c.sigla, c.nome, count(*) qtd
      from Tipo_Data c
           inner join (select a.sq_particular_calendario, b.sq_tipo_data
                         from Calendario_Cliente               b
                              inner join Particular_Calendario a on (a.sq_particular_calendario = b.sq_particular_calendario)
                        where a.sq_particular_calendario = calendario
                        and   year(b.dt_ocorrencia) = ano
                      )  d on (d.sq_tipo_data             = c.sq_tipo_data)
     where (c.sigla in ('IA','T1','I2','TA'))
    group by c.sigla, c.nome
    having count(*) > 1
    order by c.sigla, c.nome;

  -- Verifica se foi inserida alguma data que coincida com o calend�rio oficial
  cursor c_base is
    select a.sq_particular_calendario, b.dt_ocorrencia, count(*) qtd
      from Particular_Calendario           a
           inner   join Calendario_Cliente b on (a.sq_particular_calendario = b.sq_particular_calendario)
             inner join Tipo_Data          c on (b.sq_tipo_data             = c.sq_tipo_data)
             inner join Calendario_Base    d on (b.dt_ocorrencia            = d.dt_ocorrencia)
     where a.sq_particular_calendario = calendario
     and   year(b.dt_ocorrencia) = ano
    group by a.sq_particular_calendario, b.dt_ocorrencia
    order by a.sq_particular_calendario, b.dt_ocorrencia;

  -- Verifica se alguma data tem mais de uma ocorr�ncia com o mesmo tipo
  cursor c_duplicata is
    select a.sq_particular_calendario, b.dt_ocorrencia, b.sq_tipo_data, count(*) qtd
      from Particular_Calendario           a
           inner   join Calendario_Cliente b on (a.sq_particular_calendario = b.sq_particular_calendario)
             inner join Tipo_Data          c on (b.sq_tipo_data             = c.sq_tipo_data)
     where a.sq_particular_calendario = calendario
     and   year(b.dt_ocorrencia) = ano
    group by a.sq_particular_calendario, b.dt_ocorrencia, b.sq_tipo_data
    having count(*) > 1
    order by a.sq_particular_calendario, b.dt_ocorrencia;

begin

  -- Verifica se o calend�rio tem in�cio/fim de ano/semestre letivo. Qtd deve ser igual a 1
  -- Verifica se o calend�rio tem in�cio/fim de ano/semestre letivo. Qtd deve ser igual a 1
  for crec in c_ano_letivo loop
      texto := texto + '<li>' + crec.nome + ' deve ser informado';
  End Loop;

  -- Impede duplica��o nas datas de in�cio/fim de ano/semestre letivo
  for crec in c_ano_letivo_dup loop
      texto := texto + '<li>' + crec.nome + ' s� pode ter uma data indicada';
  End Loop;

  -- Verifica se foi inserida alguma data que coincida com o calend�rio oficial
  for crec in c_base loop
      texto := texto + '<li>' + to_char(crec.dt_ocorrencia,'dd/mm/yyyy') + ' consta do Calend�rio Oficial e deve ser removida deste calend�rio';
  End Loop;

  -- Verifica se alguma data tem mais de uma ocorr�ncia com o mesmo tipo
  for crec in c_duplicata loop
      texto := texto + '<li>' + to_char(crec.dt_ocorrencia,'dd/mm/yyyy') + ' s� pode ser informada uma vez';
  End Loop;

  If texto = '' Then texto := null; End If;

  Return texto;
end ValidaCalendario;
/
