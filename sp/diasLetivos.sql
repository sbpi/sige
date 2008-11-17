alter function dbo.DiasLetivos
  (@data_inicio  varchar(10),
   @data_fim     varchar(10),
   @cliente      int,
   @calendario   int = null
  )
   returns int as
/**********************************************************************************
* Nome      : DiasLetivos
* Finalidade: Calcular o nº de dias letivos entre duas datas
* Autor     : Alexandre Vinhadelli Papadópolis
* Data      :  16/09/2008, 13:00
* Parâmetros:
*    @data_inicio   : data inicial
*    @data_fim      : data final
*    @cliente       : indica o cliente
*    @calendario    : indica o calendário do cliente
* Retorno: resultado calculado para o número de dias entre as duas datas e períodos
***********************************************************************************/
begin
  Declare @w_atual    datetime;
  Declare @w_inicio   datetime;
  Declare @w_fim      datetime;
  Declare @w_dias     numeric(10,1);
  Declare @w_feriado  numeric(4,1);
  Declare @w_let_ini  datetime;
  Declare @w_let1_fim datetime;
  Declare @w_let2_ini datetime;
  Declare @w_let_fim  datetime;
  Declare @w_publica  varchar(1);
  Declare @w_ano      int;

  Set @w_inicio = convert(datetime,@data_inicio,103);
  Set @w_fim    = convert(datetime,@data_fim,103);
  Set @w_atual  = @w_inicio;

  If @w_inicio > @w_fim or year(@w_inicio) <> year(@w_fim) Begin
     -- Se o período não estiver correto, aponta o erro
     Set @w_dias = -2;
  End Else Begin
     Set @w_ano = year(@w_inicio);

     -- Verifica se é escola pública ou privada
     select @w_publica = publica
       from escCliente a
      where sq_cliente = @cliente;

     -- Verifica os inícios e términos de semestre
     If @w_publica = 'N' Begin
        select @w_let_ini = dt_ocorrencia 
          from escParticular_Calendario         c
               inner join escCalendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join escTipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'IA')
         where sq_site_cliente = @cliente
           and c.sq_particular_calendario = coalesce(@calendario,0);

        select @w_let_fim = dt_ocorrencia 
          from escParticular_Calendario         c
               inner join escCalendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join escTipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'TA')
         where sq_site_cliente = @cliente
           and c.sq_particular_calendario = coalesce(@calendario,0);

        select @w_let2_ini = dt_ocorrencia 
          from escParticular_Calendario         c
               inner join escCalendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join escTipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'I2')
         where sq_site_cliente = @cliente
           and c.sq_particular_calendario = coalesce(@calendario,0);

        select @w_let1_fim = dt_ocorrencia 
          from escParticular_Calendario         c
               inner join escCalendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join escTipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'T1')
         where sq_site_cliente = @cliente
           and c.sq_particular_calendario = coalesce(@calendario,0);
     End Else Begin
        select @w_let_ini = dt_ocorrencia 
          from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'IA')
         where sq_site_cliente = 0;

        select @w_let_fim = dt_ocorrencia 
          from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'TA')
         where sq_site_cliente = 0;

        select @w_let2_ini = dt_ocorrencia 
          from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'I2')
         where sq_site_cliente = 0;

        select @w_let1_fim = dt_ocorrencia 
          from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = @w_ano and b.sigla = 'T1')
         where sq_site_cliente = 0;
     End;

     If @w_let_ini is null or @w_let1_fim is null or @w_let2_ini is null or @w_let_fim is null Begin
        -- Se não recuperou alguma das datas, retorna erro.
        Return -2
     End;

     Set @w_dias   = 0;
     While @w_atual <= @w_fim Begin
       -- Recupera somente dentro do ano letivo
       If @w_atual between @w_let_ini and @w_let1_fim or
          @w_atual between @w_let2_ini and @w_let_fim
       Begin
          -- Trata as sextas, sábados e domingos
          If datepart(dw,@w_atual) not in (1,7) Begin
             Set @w_dias = @w_dias + 1;
          End
       End

       -- Incrementa a data atual
       Set @w_atual = @w_atual + 1;
     End;

     -- Trata os feriados oficiais
     select @w_feriado = count(*) 
       from escCalendario_Base a
      where a.dt_ocorrencia between @w_inicio and @w_fim
        and a.dia_letivo = 'N'
        and datepart(dw,a.dt_ocorrencia) not in (1,7)
        and (a.dt_ocorrencia between @w_let_ini and @w_let1_fim or
             a.dt_ocorrencia between @w_let2_ini and @w_let_fim
            );

     If @w_feriado > 0 Begin Set @w_dias = @w_dias - @w_feriado; End

     -- Trata recessos e datas da escola
     If @w_publica = 'N' Begin
        select @w_feriado = count(*) 
          from escParticular_Calendario         c
               inner join escCalendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join escTipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla in ('RE','RA','RI'))
         where a.sq_site_cliente = @cliente
           and c.sq_particular_calendario = coalesce(@calendario,0)
           and (datepart(dw,a.dt_ocorrencia) not in (1,7) or (datepart(dw,a.dt_ocorrencia) in (1,7)) and b.sigla = 'SL')
           and a.dt_ocorrencia between @w_inicio and @w_fim
           and (a.dt_ocorrencia between @w_let_ini and @w_let1_fim or
                a.dt_ocorrencia between @w_let2_ini and @w_let_fim
               );
     End Else Begin
        select @w_feriado = count(*) 
          from escCalendario_Cliente a inner join escTipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and b.sigla not in ('IA','TA','T1','I2','CN','SL','OU','PR','HC'))
         where a.sq_site_cliente = @cliente
           and (datepart(dw,a.dt_ocorrencia) not in (1,7) or (datepart(dw,a.dt_ocorrencia) in (1,7)) and b.sigla = 'SL')
           and a.dt_ocorrencia between @w_inicio and @w_fim
           and (a.dt_ocorrencia between @w_let_ini and @w_let1_fim or
                a.dt_ocorrencia between @w_let2_ini and @w_let_fim
               );
     End;

     If @w_feriado > 0 Begin Set @w_dias = @w_dias - @w_feriado; End

     -- Trata recessos e datas da regional de ensino
     select @w_feriado = count(*) 
       from escCalendario_Cliente a
            inner join escTipo_Data a1 on (a.sq_tipo_data = a1.sq_tipo_data and a1.sigla not in ('IA','TA','T1','I2','CN','SL','OU','PR','HC'))
            inner join escCliente b on (a.sq_site_cliente = b.sq_cliente)
            inner join escCliente c on (b.sq_cliente      = c.sq_cliente_pai)
      where c.sq_cliente = @cliente
        and (datepart(dw,a.dt_ocorrencia) not in (1,7) or (datepart(dw,a.dt_ocorrencia) in (1,7)) and a1.sigla = 'SL')
        and a.dt_ocorrencia between @w_inicio and @w_fim
        and (a.dt_ocorrencia between @w_let_ini and @w_let1_fim or
             a.dt_ocorrencia between @w_let2_ini and @w_let_fim
            )

     If @w_feriado > 0 Begin Set @w_dias = @w_dias - @w_feriado; End

     -- Trata recessos e datas da secretaria de educacao
     select @w_feriado = count(*) 
       from escCalendario_Cliente a
            inner join escTipo_Data a1 on (a.sq_tipo_data = a1.sq_tipo_data and a1.sigla not in ('IA','TA','T1','I2','CN','SL','OU','PR','HC'))
            inner join escCliente b on (a.sq_site_cliente = b.sq_cliente)
            inner join escCliente c on (b.sq_cliente      = c.sq_cliente_pai)
            inner join escCliente d on (c.sq_cliente      = d.sq_cliente_pai)
      where d.sq_cliente = @cliente
        and (datepart(dw,a.dt_ocorrencia) not in (1,7) or (datepart(dw,a.dt_ocorrencia) in (1,7)) and a1.sigla = 'SL')
        and a.dt_ocorrencia between @w_inicio and @w_fim
        and (a.dt_ocorrencia between @w_let_ini and @w_let1_fim or
             a.dt_ocorrencia between @w_let2_ini and @w_let_fim
            )

     If @w_feriado > 0 Begin Set @w_dias = @w_dias - @w_feriado; End
  End

  Return @w_dias;
end
